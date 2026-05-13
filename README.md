# Okto Asana Smart Sync

Upload any Excel file → AI reads it → Asana updates automatically.

## Deploy to Railway (10 minutes)

1. Go to https://github.com and create a free account if you don't have one
2. Create a new repository called `okto-sync`
3. Upload all these files to the repository
4. Go to https://railway.app and sign in with GitHub
5. Click "New Project" → "Deploy from GitHub repo" → select `okto-sync`
6. Once deployed, go to "Variables" and add:
   ```
   ANTHROPIC_API_KEY = your_anthropic_api_key_here
   ```
7. Railway gives you a public URL — share it with your team

## That's it

Everyone in the office goes to the URL, uploads their Excel, enters their own Asana token, clicks sync.

## Local development

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=your_key_here
python app.py
```

Then open http://localhost:5000
