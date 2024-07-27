import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../auth/authConfig";


function Login() {
  const { instance } = useMsal();

  const handleLogin = () => {
    instance.loginPopup(loginRequest).catch((e) => {
      console.error(e);
    });
  };

  return <button onClick={handleLogin}>Sign in with Microsoft</button>;
}

export default Login;
