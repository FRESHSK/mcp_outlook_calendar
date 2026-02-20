# MCP Outlook Server

Un serveur **Model Context Protocol (MCP)** pour Microsoft Outlook, permettant à des agents IA comme **Claude Desktop** d'interagir avec ta boîte mail et ton calendrier via l'API Microsoft Graph.

---

## ✨ Fonctionnalités

- 📧 **Mail** : Lire, consulter et envoyer des emails
- 📅 **Calendrier** : Lister et créer des événements avec lien Teams automatique
- 🔐 **Authentification** : Azure AD Device Code Flow avec cache de token persistant
- 🔄 **Auto-refresh** : Le token se renouvelle automatiquement

---

## 🛠️ Outils disponibles

| Outil | Description | Paramètres |
|-------|-------------|------------|
| `list_messages` | Liste les 10 derniers emails de la boîte de réception | — |
| `get_message` | Lit le contenu complet d'un email | `message_id` |
| `send_message` | Envoie un email | `subject`, `toRecipients`, `content`, `contentType` |
| `list_events` | Liste les événements des 7 prochains jours | — |
| `create_event` | Crée un événement (avec lien Teams automatique) | `subject`, `startDateTime`, `endDateTime`, `location`, `content`, `attendees` |

---

## 🧰 Prérequis

- [Node.js 18+](https://nodejs.org)
- Un compte Microsoft (personnel ou professionnel)

---

## 🚀 Installation étape par étape

### Étape 1 — Cloner le projet

```bash
git clone https://github.com/FRESHSK/mcp_outlook_01.git
cd mcp_outlook_01
npm install
```

---

### Étape 2 — Créer une App Azure AD (une seule fois)

1. Aller sur [Azure Portal - App Registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps)
2. Cliquer **"New registration"**
3. Remplir le formulaire :
   - **Name** : `MCP Outlook`
   - **Supported account types** : choisir la 3ème option *(Multitenant + personal Microsoft accounts)*
   - **Redirect URI** : `Public client/native` → `http://localhost`
   - Cliquer **Register**
4. Activer les flux publics :
   - Menu gauche → **Authentication**
   - Scroll bas → **Advanced settings**
   - **Allow public client flows** → **Yes**
   - Cliquer **Save**
5. Copier l'**Application (client) ID** depuis **Overview**

---

### Étape 3 — Ajouter les permissions Microsoft Graph API

1. Dans le [Azure Portal](https://portal.azure.com), aller sur ton App Registration
2. Menu gauche → **API permissions**
3. Cliquer **"Add a permission"**
4. Choisir **Microsoft Graph**
5. Choisir **Delegated permissions**
6. Rechercher et cocher les permissions suivantes :

| Permission | Type | Description |
|-----------|------|-------------|
| `User.Read` | Delegated | Lire le profil de l'utilisateur connecté |
| `Mail.Read` | Delegated | Lire les emails |
| `Mail.Read.Shared` | Delegated | Lire les emails partagés |
| `Mail.ReadBasic` | Delegated | Lire les infos de base des emails |
| `Mail.ReadWrite` | Delegated | Lire et modifier les emails |
| `Mail.ReadWrite.Shared` | Delegated | Lire et modifier les emails partagés |
| `Mail.Send` | Delegated | Envoyer des emails |
| `Mail.Send.Shared` | Delegated | Envoyer des emails au nom d'autres |
| `Calendars.Read` | Delegated | Lire les calendriers |
| `Calendars.Read.Shared` | Delegated | Lire les calendriers partagés |
| `Calendars.ReadBasic` | Delegated | Lire les infos de base des calendriers |
| `Calendars.ReadWrite` | Delegated | Lire et modifier les calendriers |
| `Calendars.ReadWrite.Shared` | Delegated | Lire et modifier les calendriers partagés |

7. Cliquer **"Add permissions"**
8. Cliquer **"Grant admin consent for ..."** puis confirmer

**Résultat attendu dans Azure Portal :**

```
✅ User.Read                    (Delegated) - Granted
✅ Mail.Read                    (Delegated) - Granted
✅ Mail.Read.Shared             (Delegated) - Granted
✅ Mail.ReadBasic               (Delegated) - Granted
✅ Mail.ReadWrite               (Delegated) - Granted
✅ Mail.ReadWrite.Shared        (Delegated) - Granted
✅ Mail.Send                    (Delegated) - Granted
✅ Mail.Send.Shared             (Delegated) - Granted
✅ Calendars.Read               (Delegated) - Granted
✅ Calendars.Read.Shared        (Delegated) - Granted
✅ Calendars.ReadBasic          (Delegated) - Granted
✅ Calendars.ReadWrite          (Delegated) - Granted
✅ Calendars.ReadWrite.Shared   (Delegated) - Granted
```

---

### Étape 4 — Configurer les variables d'environnement

```bash
cp .env.example .env
```

Éditer `.env` :

```env
MICROSOFT_CLIENT_ID=ton-client-id-azure
MICROSOFT_TENANT_ID=common
```

> ⚠️ Ne jamais committer le fichier `.env` — il est dans `.gitignore`

> ℹ️ Utilise `common` pour un compte personnel, ou ton Tenant ID pour un compte organisation.

---

### Étape 5 — Première authentification Microsoft

```bash
node server.js
```

Tu verras dans le terminal :
```
To sign in, use a web browser to open the page
https://microsoft.com/devicelogin and enter the code XXXXXXXX
```

1. Ouvrir [https://microsoft.com/devicelogin](https://microsoft.com/devicelogin)
2. Entrer le code affiché dans le terminal
3. Se connecter avec ton compte Microsoft
4. Un fichier `token_cache.json` sera créé automatiquement

> ✅ Cette étape n'est nécessaire qu'une seule fois. Le token se renouvelle automatiquement.

---

### Étape 6 — Configurer Claude Desktop

Éditer le fichier `claude_desktop_config.json` de Claude Desktop :
- **Windows** : `C:\Users\<user>\AppData\Roaming\Claude\claude_desktop_config.json`
- **Mac** : `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": [
        "C:\\chemin\\absolu\\vers\\mcp_outlook_01\\server.js"
      ]
    }
  }
}
```

---

### Étape 7 — Redémarrer Claude Desktop

Fermer et relancer Claude Desktop. Tu peux maintenant demander à Claude :

- *"Montre-moi mes derniers emails"*
- *"Lis l'email de GitHub"*
- *"Envoie un email à mon collègue"*
- *"Qu'est-ce que j'ai de prévu cette semaine ?"*
- *"Planifie une réunion avec [email] demain à 10h"* (crée automatiquement un lien Teams)

---

## 📁 Structure du projet

```
mcp_outlook_01/
├── server.js              # Serveur MCP principal
├── .env                   # Variables d'environnement (non commité)
├── .env.example           # Template des variables d'environnement
├── .gitignore             # Fichiers ignorés par Git
├── token_cache.json       # Cache du token MSAL (non commité)
└── package.json           # Dépendances Node.js
```

---

## 📦 Dépendances principales

| Package | Usage |
|---------|-------|
| `@modelcontextprotocol/sdk` | Serveur MCP |
| `@azure/msal-node` | Authentification Microsoft |
| `node-fetch` | Requêtes HTTP vers Graph API |
| `dotenv` | Variables d'environnement |
| `zod` | Validation des paramètres des outils |

---

## 🔒 Sécurité

- Le fichier `.env` contenant les clés est dans `.gitignore`
- Le `token_cache.json` contenant les tokens Microsoft est dans `.gitignore`
- Ne jamais committer de clés API ou de tokens dans GitHub

---

## 🐛 Dépannage

**Le MCP ne répond pas dans Claude Desktop ?**
→ Vérifier les logs : `AppData\Roaming\Claude\logs\mcp-server-outlook.log`

**Erreur d'authentification ou nouvelles permissions ?**
→ Supprimer `token_cache.json` et relancer `node server.js` pour se réauthentifier

**Erreur `fetch is not a function` ?**
→ Vérifier que `node-fetch` est bien installé : `npm install node-fetch`

**Erreur lors de l'envoi d'email ?**
→ Le code gère les réponses 202 Accepted — vérifier que `server.js` est à jour

---

## 📝 Licence

MIT
