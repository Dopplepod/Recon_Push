# P&L Reconciliation V12

This build now includes:
- Login screen
- Session-based authentication
- Docker support for Render deployment
- `render.yaml` Blueprint for one-click setup from a Git repo

## Default login
- Username: `finance`
- Password: `amcosg`

## Run locally without Docker
```bash
pip install -r requirements.txt
python app.py
```
Then open `http://127.0.0.1:5000`

## Run locally with Docker
```bash
docker build -t recon-v12 .
docker run -p 10000:10000 recon-v12
```
Then open `http://127.0.0.1:10000`

## Deploy to Render
1. Upload this project to a GitHub repo.
2. In Render, choose **New +** → **Blueprint**.
3. Select your repo.
4. Render will detect `render.yaml` and create the web service.
5. After deploy, open the generated URL and log in.

## Important Render notes
- The container serves the app with Gunicorn.
- The Flask secret key can be overridden with the `SECRET_KEY` environment variable.
- Health check path is `/api/session`.

## Included reference files
- `data/mappings/BFC_To_OS_Mapping.xlsx`
- `data/reference/hierarchy.xml`
