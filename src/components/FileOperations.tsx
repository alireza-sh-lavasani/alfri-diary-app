import React, { useState, useEffect, useRef } from "react";
import { useMsal } from "@azure/msal-react";
import { Client } from "@microsoft/microsoft-graph-client";
import axios from "axios";
import { loginRequest } from "../auth/authConfig";
import dayjs from "dayjs";
import { DatePicker } from "@mui/x-date-pickers";
import {
  Button,
  Box,
  Typography,
  CircularProgress,
  LinearProgress,
  Stack,
  Grid,
  Fab,
  Avatar,
} from "@mui/material";
import { CloudUpload, Create, Description } from "@mui/icons-material";
interface Thumbnail {
  url: string;
  width: number;
  height: number;
}

interface ThumbnailSet {
  small?: Thumbnail;
  medium?: Thumbnail;
  large?: Thumbnail;
}

interface Document {
  id: string;
  name: string;
  webUrl: string;
  createdDateTime: string;
  thumbnails?: ThumbnailSet[];
}

interface GroupedDocuments {
  [year: string]: Document[];
}

// Mock data generator
const generateMockData = (selectedDate: dayjs.Dayjs): Document[] => {
  const currentDate = selectedDate;
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
  const [selectedDate, setSelectedDate] = useState<dayjs.Dayjs>(dayjs());

  useEffect(() => {
    listDocuments();
  }, [selectedDate]);

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
    const now = selectedDate;
    const currentYear = now.year();
    const month = now.format("MM");
    const day = now.format("DD");

    try {
      let allDocuments: Document[] = [];

      if (process.env.REACT_APP_USE_MOCK_DOCUMENTS === "true") {
        allDocuments = generateMockData(selectedDate);
      } else {
        setIsLoading(true);

        // Fetch documents for the last 10 years (you can adjust this number)
        for (let year = currentYear; year > currentYear - 10; year--) {
          const startDate = dayjs(`${year}-${month}-${day}`)
            .startOf("day")
            .toISOString();
          const endDate = dayjs(startDate).add(1, "day").toISOString();

          const result = await graphClient
            .api("/me/drive/root:/Documents:/children")
            .filter(
              `createdDateTime ge ${startDate} and createdDateTime lt ${endDate}`
            )
            .select("id,name,webUrl,createdDateTime,thumbnails")
            .orderby("createdDateTime desc")
            .top(1000) // Adjust this number based on your needs
            .expand("thumbnails")
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
      setIsLoading(false);
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
    <Box sx={{ padding: 2 }}>
      <Stack>
        <DatePicker
          value={selectedDate}
          onChange={(value) => setSelectedDate(value || dayjs())}
          disableFuture
        />

        <Grid container gap={1} sx={{ mt: 3 }}>
          <Fab color="primary" onClick={handleCreateAndOpen}>
            <Create />
          </Fab>

          <Fab
            color="secondary"
            onClick={() => fileInputRef.current?.click()}
            disabled={isUploading}
          >
            <CloudUpload />

            <input
              type="file"
              ref={fileInputRef}
              style={{ display: "none" }}
              onChange={handleFileUpload}
              disabled={isUploading}
            />
          </Fab>
        </Grid>
      </Stack>

      {isUploading && (
        <Box sx={{ width: "100%", mt: 2 }}>
          <LinearProgress variant="determinate" value={uploadProgress} />
        </Box>
      )}

      <Typography variant="h4" sx={{ mt: 4 }}>
        On This Day:
      </Typography>

      {isLoading ? (
        <Box sx={{ display: "flex", justifyContent: "center", mt: 2 }}>
          <CircularProgress />
        </Box>
      ) : (
        Object.entries(groupedDocuments)
          .sort(([a], [b]) => parseInt(b) - parseInt(a))
          .map(([year, docs]) => (
            <Box key={year} sx={{ mt: 3 }}>
              <Typography variant="h5">
                {getYearLabel(year)} ({year})
              </Typography>
              <Box component="ul" sx={{ paddingLeft: 2 }}>
                {docs.map((doc) => {
                  if (doc.name.slice(-5) === ".docx")
                    return (
                      <Box
                        component="li"
                        key={doc.id}
                        sx={{ mt: 1, listStyleType: "none" }}
                        >
                        <Button
                          variant="outlined"
                          color="primary"
                          sx={{
                            display: "flex",
                            alignItems: "center",
                            gap: "8px",
                            padding: "5px 10px",
                          }}
                          onClick={() => openDocument(doc.webUrl)}
                        >
                          <Fab size="small" color="primary">
                            <Description />
                          </Fab>
                          {doc.name}
                        </Button>
                      </Box>
                    );
                  else if (doc?.thumbnails && doc?.thumbnails.length > 0)
                    return (
                      <Box
                        component="li"
                        key={doc.id}
                        sx={{ mt: 1, listStyleType: "none" }}
                      >
                        <Button
                          variant="outlined"
                          sx={{
                            display: "flex",
                            alignItems: "center",
                            gap: "8px",
                            padding: "5px 10px",
                          }}
                          onClick={() => openDocument(doc.webUrl)}
                        >
                          <Avatar
                            // @ts-ignore
                            src={doc.thumbnails[0].small.url}
                            alt="Thumbnail"
                          />
                          {doc.name}
                        </Button>
                      </Box>
                    );
                  else
                    <Box
                      component="li"
                      key={doc.id}
                      sx={{ mt: 1, listStyleType: "none" }}
                    >
                      <Button
                        onClick={() => openDocument(doc.webUrl)}
                        variant="outlined"
                      >
                        {doc.name}
                      </Button>
                    </Box>;
                })}
              </Box>
            </Box>
          ))
      )}
    </Box>
  );
};

export default FileOperations;
