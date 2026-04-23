import React, { useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./authConfig";

const msalInstance = new PublicClientApplication(msalConfig);

function App() {
  const [data, setData] = useState({ Naam: "", Email: "", Bericht: "" });

  const handleChange = (e) => {
    setData({ ...data, [e.target.name]: e.target.value });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    try {
      const account = msalInstance.getAllAccounts()[0];
      let accessToken;
      if (!account) {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        accessToken = loginResponse.accessToken;
      } else {
        const tokenResponse = await msalInstance.acquireTokenSilent({
          ...loginRequest,
          account
        });
        accessToken = tokenResponse.accessToken;
      }

      const response = await fetch(
        "[graph.microsoft.com](https://graph.microsoft.com/v1.0/sites/)<SITE_ID>/lists/<LIST_ID>/items",
        {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            fields: {
              Naam: data.Naam,
              Email: data.Email,
              Bericht: data.Bericht
            }
          })
        }
      );

      if (response.ok) {
        alert("Inzending opgeslagen!");
        setData({ Naam: "", Email: "", Bericht: "" });
      } else {
        alert("Er ging iets mis bij opslaan.");
        console.log(await response.text());
      }
    } catch (err) {
      console.error(err);
    }
  };

  return (
    <div style={{ padding: "20px" }}>
      <h2>React formulier → SharePoint</h2>
      <form onSubmit={handleSubmit}>
        <input name="Naam" value={data.Naam} onChange={handleChange} placeholder="Naam" /><br></br>
        <textarea name="Bericht" value={data.Bericht} onChange={handleChange} placeholder="Bericht" /><br></br>
        <button type="submit">Verzenden</button>
      </form>
    </div>
   
  );
}

export default App;
