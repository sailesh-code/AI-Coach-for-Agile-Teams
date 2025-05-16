# Sprint Report Generator

A web application that generates detailed sprint reports by integrating with Jira and using Google's Gemini AI to break down sprint goals into relevant subgoals.

## Features

- Fetches data from Jira API for the last closed sprint
- Uses Gemini AI to generate relevant subgoals from sprint goals
- Modern React frontend with Material-UI
- Flask backend with CORS support

## Prerequisites

- Python 3.7+
- Node.js 14+
- Jira account with API access
- Google Gemini API key

## Setup

1. Clone the repository
2. Set up the backend:
   ```bash
   # Create and activate virtual environment (optional but recommended)
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate

   # Install dependencies
   pip install -r requirements.txt

   # Create .env file from template
   cp .env.example .env
   ```

3. Set up the frontend:
   ```bash
   # Install dependencies
   npm install
   ```

4. Configure environment variables in `.env`:
   - `JIRA_URL`: Your Jira instance URL
   - `JIRA_EMAIL`: Your Jira email
   - `JIRA_API_TOKEN`: Your Jira API token
   - `GEMINI_API_KEY`: Your Google Gemini API key

## Running the Application

1. Start the backend server:
   ```bash
   python app.py
   ```

2. Start the frontend development server:
   ```bash
   npm start
   ```

3. Open your browser and navigate to `http://localhost:3000`

## Usage

1. Click the "Generate Report" button on the homepage
2. The application will fetch the last closed sprint from Jira
3. The sprint goal will be processed by Gemini AI to generate relevant subgoals
4. The complete report will be displayed on the page

## API Endpoints

- `GET /api/sprint-report`: Fetches the last closed sprint report with AI-generated subgoals

## Technologies Used

- Frontend: React, Material-UI
- Backend: Python Flask
- AI: Google Gemini
- Integration: Jira API 