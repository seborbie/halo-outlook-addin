# Halo Outlook Add-in

The Outlook add-in lives in `Halo Email Integration`.

## Local Halo OAuth setup

Register a Halo API application with the `Authorisation Code (Native Application)` authentication method and add this redirect URI:

```text
https://localhost:3000/auth/callback
```

Start the Office add-in dev server:

```powershell
cd ".\Halo Email Integration"
npm run dev-server
```

On first use, the task pane asks for:

- Halo URL, for example `https://your-company.halopsa.com`
- Halo API application client ID

The client ID is public OAuth application metadata. The add-in does not ask for or use a client secret.
