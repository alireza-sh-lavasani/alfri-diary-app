import React, { useState, useEffect, useRef } from "react";
import { useMsal } from "@azure/msal-react";
import { Client } from "@microsoft/microsoft-graph-client";
import axios from "axios";
import { loginRequest } from "../auth/authConfig";
import dayjs from "dayjs";

interface Document {
  id: string;
  name: string;
  webUrl: string;
  createdDateTime: string;
}

const FileOperations: React.FC = () => {
  const { instance, accounts } = useMsal();
  const [documents, setDocuments] = useState<Document[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    listDocuments();
  }, []);

  const getGraphClient = async () => {
    const account = accounts[0];
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account: account,
    });
    return Client.init({
      authProvider: (done) => {
        done(null, response.accessToken);
      },
    });
  };

  const listDocuments = async () => {
    const graphClient = await getGraphClient();
    const today = dayjs().startOf("day").toISOString();
    const result = await graphClient
      .api("/me/drive/root:/Documents:/children")
      .filter(`createdDateTime ge ${today}`)
      .select("id,name,webUrl,createdDateTime")
      .orderby("createdDateTime desc")
      .get();
    setDocuments(result.value);
  };

  const createDocument = async () => {
    const graphClient = await getGraphClient();
    const baseFileName = `NewDocument_${dayjs().format("YYYY-MM-DD")}`;
    const emptyContent = new Blob([""], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    let fileName = `${baseFileName}.docx`;
    let index = 1;

    while (true) {
      try {
        await graphClient.api(`/me/drive/root:/Documents/${fileName}`).get();
        fileName = `${baseFileName}(${index}).docx`;
        index++;
      } catch (error) {
        break;
      }
    }

    try {
      const response = await graphClient
        .api(`/me/drive/root:/Documents/${fileName}:/content`)
        .put(emptyContent);

      await listDocuments();
      return response.webUrl;
    } catch (error) {
      console.error("Error creating document:", error);
    }
  };

  const uploadFile = async (file: File) => {
    setIsUploading(true);
    setUploadProgress(0);

    try {
      const totalSize = file.size;
      const uploadUrl = `/me/drive/root:/Documents/${file.name}:/content`;

      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      await axios.put(
        `https://graph.microsoft.com/v1.0${uploadUrl}`,
        file,
        {
          headers: {
            "Content-Type": file.type,
            Authorization: `Bearer ${tokenResponse.accessToken}`,
          },
          onUploadProgress: (progressEvent) => {
            const progress = Math.round((progressEvent.loaded / totalSize) * 100);
            setUploadProgress(progress);
          },
        }
      );

      await listDocuments();
    } catch (error) {
      console.error("Error uploading file:", error);
    } finally {
      setIsUploading(false);
      setUploadProgress(0);
    }
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      uploadFile(file);
    }
  };

  const openDocument = (webUrl: string) => {
    window.open(webUrl, "_blank");
  };

  const handleCreateAndOpen = async () => {
    const newDocumentUrl = await createDocument();
    if (newDocumentUrl) {
      openDocument(newDocumentUrl);
    }
  };

  return (
    <div>
      <button onClick={handleCreateAndOpen}>
        Create and Open New Word Document
      </button>
      <input
        type="file"
        ref={fileInputRef}
        style={{ display: "none" }}
        onChange={handleFileUpload}
        disabled={isUploading}
      />
      <button
        onClick={() => fileInputRef.current?.click()}
        disabled={isUploading}
      >
        {isUploading ? `Uploading... ${uploadProgress}%` : "Upload File"}
      </button>
      {isUploading && (
        <div
          style={{
            width: "200px",
            backgroundColor: "#e0e0e0",
            marginTop: "10px",
          }}
        >
          <div
            style={{
              width: `${uploadProgress}%`,
              backgroundColor: "#4CAF50",
              height: "20px",
              transition: "width 0.3s ease-in-out",
            }}
          />
        </div>
      )}
      <h2>Your Documents Created Today:</h2>
      <ul>
        {documents.map((doc) => (
          <li key={doc.id}>
            {doc.name}
            <button onClick={() => openDocument(doc.webUrl)}>Open</button>
          </li>
        ))}
      </ul>
    </div>
  );
};

export default FileOperations;