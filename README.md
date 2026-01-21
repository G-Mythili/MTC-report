# MTC Report Studio

A professional web-based tool for analyzing Spectro reports and generating MTC (Material Test Certificate) documents.

## Project Structure

- `frontend/`: React + Vite application (Dashboard UI)
- `backend/`: FastAPI application (Excel processing logic)

## Deployment / Cloud Setup

### 1. GitHub Codespaces (Recommended)
1. Click the **Code** button on GitHub and select **Codespaces** -> **Create codespace**.
2. Once the environment opens:
   - **Backend**: Open a terminal, run:
     ```bash
     cd backend
     pip install -r requirements.txt
     uvicorn main:app --host 0.0.0.0 --port 8000
     ```
   - **Frontend**: Open a second terminal, run:
     ```bash
     cd frontend
     npm install
     npm run dev -- --host
     ```
3. GitHub will provide a link to the frontend. Ensure port **8000** is set to **Public** in the Ports tab so the frontend can reach the API.

### 2. Local Setup
See the `RUN_MTC_STUDIO.bat` for quick local startup.

## Technologies Used
- **Frontend**: React, Tailwind CSS, Lucide icons, Vite
- **Backend**: Python, FastAPI, Pandas, Openpyxl
