---
name: mcp-setup
description: Guida alla configurazione e autorizzazione del server MCP Google Docs
---

# MCP Setup: Google Docs

Questo server richiede l'accesso alle API di Google (Docs, Sheets, Drive).

## Configurazione Credenziali

1. **credentials.json**: Deve essere presente nella root del progetto.
   - Deve contenere un oggetto `installed` scaricato dalla Google Cloud Console (Desktop App).
   - Formato atteso:
     ```json
     {
       "installed": {
         "client_id": "...",
         "project_id": "...",
         "auth_uri": "https://accounts.google.com/o/oauth2/auth",
         "token_uri": "https://oauth2.googleapis.com/token",
         "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
         "client_secret": "...",
         "redirect_uris": ["http://localhost"]
       }
     }
     ```

2. **Autorizzazione**:
   - Per generare il token di accesso iniziale, eseguire:
     ```bash
     npm start auth
     ```
   - Questo aprirà il browser per l'OAuth flow.
   - Il token verrà salvato in `~/.config/google-docs-mcp/token.json`.

## Risoluzione Problemi

- **"Could not find client secrets in credentials.json"**: Verificare che il file sia nel formato `installed` (non `web` o con `type: authorized_user` al primo livello).
- **"Server failed to start: Google client initialization failed"**: Spesso dovuto a un token scaduto o invalido. Rilanciare `npm start auth`.
