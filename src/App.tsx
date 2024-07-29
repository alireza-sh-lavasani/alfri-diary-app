// src/App.tsx
import { useState, useEffect } from "react";
import {
  BrowserRouter as Router,
  Route,
  Routes,
  Navigate,
} from "react-router-dom";
import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
  useMsal,
} from "@azure/msal-react";
import {
  AppBar,
  Toolbar,
  Typography,
  Avatar,
  Container,
  CssBaseline,
  Box,
} from "@mui/material";
import Login from "./components/Login";
import Logout from "./components/Logout";
import FileOperations from "./components/FileOperations";
import { LocalizationProvider } from "@mui/x-date-pickers";
import { AdapterDayjs } from "@mui/x-date-pickers/AdapterDayjs";
import { getProfilePhoto } from "./graphService";

const App: React.FC = () => {
  const { instance, accounts } = useMsal();
  const [profilePhoto, setProfilePhoto] = useState<string | null>(null);

  useEffect(() => {
    if (accounts.length > 0) {
      getProfilePhoto(instance, accounts[0]).then((photo) => {
        setProfilePhoto(photo);
      });
    }
  }, [instance, accounts]);

  const userName = accounts[0]?.name || "";

  return (
    <LocalizationProvider dateAdapter={AdapterDayjs}>
      <Router>
        <>
          <CssBaseline />
          <AppBar position="static">
            <Toolbar>
              <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
                Diary App
              </Typography>
              {accounts.length > 0 && (
                <Box display="flex" alignItems="center">
                  <Typography variant="body1" sx={{ mr: 2 }}>
                    {userName}
                  </Typography>
                  <Avatar src={profilePhoto || undefined} alt={userName} />
                </Box>
              )}
            </Toolbar>
          </AppBar>
          <Container>
            <AuthenticatedTemplate>
              <Routes>
                <Route path="/files" element={<FileOperations />} />
                <Route path="/" element={<Navigate to="/files" />} />
              </Routes>
              <Logout />
            </AuthenticatedTemplate>
            <UnauthenticatedTemplate>
              <Routes>
                <Route path="/" element={<Login />} />
                <Route path="*" element={<Navigate to="/" />} />
              </Routes>
            </UnauthenticatedTemplate>
          </Container>
        </>
      </Router>
    </LocalizationProvider>
  );
};

export default App;
