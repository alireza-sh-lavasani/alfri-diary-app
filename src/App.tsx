import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
  useIsAuthenticated,
  useMsal,
} from "@azure/msal-react";
import Login from "./components/Login";

console.log(process.env);

function App() {
  const { accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  return (
    <div>
      <h1>Diary App</h1>
      <AuthenticatedTemplate>
        {isAuthenticated && <b>Authenticated</b>}
        <p>Welcome, {accounts[0]?.name}</p>
        {/* <FileOperations /> */}
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <Login />
      </UnauthenticatedTemplate>
    </div>
  );
}

export default App;
