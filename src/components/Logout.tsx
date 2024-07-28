import { useMsal } from "@azure/msal-react";
import { useNavigate } from 'react-router-dom';

function Logout() {
  const { instance } = useMsal();
  const navigate = useNavigate();

  const handleLogout = async () => {
    await instance.logoutPopup({
      postLogoutRedirectUri: "/",
      mainWindowRedirectUri: "/"
    });
    navigate('/');
  }

  return (
    <button onClick={handleLogout}>Logout</button>
  );
}

export default Logout;