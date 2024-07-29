import React from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../auth/authConfig";
import {
  Container,
  Box,
  Typography,
  Button,
  CssBaseline,
  Paper,
} from "@mui/material";

function Login() {
  const { instance } = useMsal();

  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch((e) => {
      console.error(e);
    });
  };

  return (
    <Container component="main" maxWidth="xs">
      <CssBaseline />
      <Paper
        elevation={3}
        sx={{
          marginTop: 8,
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          padding: 2,
        }}
      >
        <Typography component="h1" variant="h4">
          Diary App
        </Typography>
        <Typography component="h2" variant="h5" sx={{ mt: 2 }}>
          Sign in with Microsoft
        </Typography>
        <Box mt={2}>
          <Button
            type="button"
            fullWidth
            variant="contained"
            color="primary"
            onClick={handleLogin}
          >
            Sign in
          </Button>
        </Box>
      </Paper>
    </Container>
  );
}

export default Login;
