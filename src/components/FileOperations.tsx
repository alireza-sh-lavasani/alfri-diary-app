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

interface GroupedDocuments {
  [year: string]: Document[];
}

// Mock data generator
const generateMockData = (): Document[] => {
  const currentDate = dayjs();
  const mockData: Document[] = [];

  for (let year = currentDate.year(); year >= currentDate.year() - 5; year--) {
    const docsCount = Math.floor(Math.random() * 5) + 1; // 1 to 5 documents per year
    for (let i = 0; i < docsCount; i++) {
      mockData.push({
        id: `doc-${year}-${i}`,
        name: `Document ${i + 1} of ${year}`,
        webUrl: `https://example.com/doc-${year}-${i}`,
        createdDateTime: dayjs(
          `${year}-${currentDate.format("MM-DD")}T12:00:00Z`
        ).toISOString(),
      });
    }
  }

  return mockData;
};

const FileOperations: React.FC = () => {
  const { instance, accounts } = useMsal();
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [groupedDocuments, setGroupedDocuments] = useState<GroupedDocuments>(
    {}
  );
  const [isLoading, setIsLoading] = useState<boolean>(false);

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
    const now = dayjs();
    const currentYear = now.year();
    const month = now.format("MM");
    const day = now.format("DD");

    try {
      let allDocuments: Document[] = [];

      if (process.env.REACT_APP_USE_MOCK_DOCUMENTS === "true") {
        allDocuments = generateMockData();
      } else {
        setIsLoading(true);

        // Fetch documents for the last 10 years (you can adjust this number)
        for (let year = currentYear; year > currentYear - 10; year--) {
          const startDate = `${year}-${month}-${day}T00:00:00Z`;
          const endDate = `${year}-${month}-${parseInt(day) + 1}T00:00:00Z`;

          const result = await graphClient
            .api("/me/drive/root:/Documents:/children")
            .filter(
              `createdDateTime ge ${startDate} and createdDateTime lt ${endDate}`
            )
            .select("id,name,webUrl,createdDateTime")
            .orderby("createdDateTime desc")
            .top(1000) // Adjust this number based on your needs
            .get();

          allDocuments = [...allDocuments, ...result.value];
        }
      }

      const grouped = allDocuments.reduce(
        (acc: GroupedDocuments, doc: Document) => {
          const year = dayjs(doc.createdDateTime).format("YYYY");
          if (!acc[year]) {
            acc[year] = [];
          }
          acc[year].push(doc);
          return acc;
        },
        {}
      );

      setGroupedDocuments(grouped);
      setIsLoading(false);
    } catch (error) {
      console.error("Error fetching documents:", error);
    }
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

      await axios.put(`https://graph.microsoft.com/v1.0${uploadUrl}`, file, {
        headers: {
          "Content-Type": file.type,
          Authorization: `Bearer ${tokenResponse.accessToken}`,
        },
        onUploadProgress: (progressEvent) => {
          const progress = Math.round((progressEvent.loaded / totalSize) * 100);
          setUploadProgress(progress);
        },
      });

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

  const getYearLabel = (year: string): string => {
    const currentYear = dayjs().year();
    const yearDiff = currentYear - parseInt(year);

    switch (yearDiff) {
      case 0:
        return "This year";
      case 1:
        return "1 year ago";
      default:
        return `${yearDiff} years ago`;
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
      <h2>On This Day:</h2>
      {isLoading ? (
        <p>Loading ...</p>
      ) : (
        Object.entries(groupedDocuments)
          .sort(([a], [b]) => parseInt(b) - parseInt(a))
          .map(([year, docs]) => (
            <div key={year}>
              <h3>
                {getYearLabel(year)} ({year})
              </h3>
              <ul>
                {docs.map((doc) => (
                  <li key={doc.id}>
                    {doc.name}
                    <button onClick={() => openDocument(doc.webUrl)}>
                      Open
                    </button>
                  </li>
                ))}
              </ul>
            </div>
          ))
      )}
    </div>
  );
};

export default FileOperations;
