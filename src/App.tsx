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
import Login from "./components/Login";
import Logout from "./components/Logout";
import FileOperations from "./components/FileOperations";

function App() {
  const { accounts } = useMsal();

  return (
    <Router>
      <div>
        <h1>Diary App</h1>
        <AuthenticatedTemplate>
          <p>Welcome, {accounts[0]?.name}</p>
          <Logout />
          <Routes>
            <Route path="/files" element={<FileOperations />} />
            <Route path="/" element={<Navigate to="/files" />} />
          </Routes>
        </AuthenticatedTemplate>
        <UnauthenticatedTemplate>
          <Routes>
            <Route path="/" element={<Login />} />
            <Route path="*" element={<Navigate to="/" />} />
          </Routes>
        </UnauthenticatedTemplate>
      </div>
    </Router>
  );
}

export default App;
