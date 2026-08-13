# VisionBI - AI Data Analyst

VisionBI is an intelligent AI-powered data analyst tool designed to quickly transform raw data files into interactive dashboards with AI-generated insights and recommendations. It automates the entire data analysis pipeline, from file ingestion and data profiling to KPI computation, visualization selection, and natural language interaction.

## Features

-   **Multi-format Data Upload**: Supports CSV, TSV, XLSX, and XLS file formats.
-   **Automated Data Pipeline**:
    -   File type detection
    -   Dataset reading
    -   Data cleaning & validation
    -   Schema detection
    -   Business context understanding
    -   Relationship discovery
    -   KPI computation
    -   Visualization selection
    -   AI insight generation
-   **Interactive Dashboard**: Displays key performance indicators (KPIs), trends, category breakdowns, distributions, and a data profile.
-   **AI-Powered Insights**: Generates executive summaries, key insights, and actionable recommendations using the Anthropic Claude API.
-   **Chat with Your Data**: An embedded AI chat interface allows users to ask natural language questions about their dataset. Includes a fallback offline mode if the AI API key is not configured.
-   **Data Quality Scoring**: Provides a data quality score based on completeness and duplicate detection.
-   **Dynamic Theming**: Toggle between dark and light themes.
-   **Report Export**: Export a summary of the analysis as a Markdown file.
-   **Sample Data**: Option to load a sample retail dataset for quick demonstration.

## Technologies Used

-   **Frontend**: React.js
-   **Charting**: Recharts
-   **Data Parsing**: PapaParse (CSV/TSV), XLSX (Excel)
-   **Icons**: Lucide-React
-   **AI Integration**: Anthropic Claude API
-   **Build Tool**: Vite
-   **Runtime**: Node.js

## Setup and Installation

To get VisionBI up and running on your local machine, follow these steps:

1.  **Prerequisites**:
    *   **Node.js**: Ensure Node.js is installed (version 18 or higher is recommended). You can download it from [nodejs.org](https://nodejs.org/).

2.  **Clone the Repository** (if applicable, assuming this is a project folder):
   ```bash
   # If this project is in a git repository
   git clone <repository-url>
   cd VisionBi
   ```
   If you already have the folder, navigate into it:
   ```bash
   cd c:\Users\PADMA\OneDrive\Desktop\ML projects\VisionBi
   ```

3.  **Install Dependencies**:
    The `run.bat` script will automatically check for and install dependencies on the first run.

4.  **Configure AI (Optional but Recommended)**:
    To enable live AI insights and chat, you need an Anthropic API key.
    *   Create a `.env` file in the project root (if it doesn't exist, `run.bat` will create one from `.env.example`).
    *   Add your Anthropic API key to this file:
        ```
        VITE_ANTHROPIC_API_KEY=your_anthropic_api_key_here
        ```
    *   You can obtain an API key from the Anthropic Console.

## Usage

To start the VisionBI application, simply run the provided batch file:

```bash
run.bat
```

This script will:
-   Check for Node.js.
-   Install npm dependencies if `node_modules` is not found.
-   Create a `.env` file from `.env.example` if it doesn't exist.
-   Open the application in your default web browser at `http://localhost:5173`.
-   Start the development server.

Press `Ctrl+C` in the command window to stop the server.

## Future Enhancements (from TODO.md)

-   Add forecasting pipeline + model registry
-   Add exports (PDF/PPT/Excel) and scheduled reports (phase 2)
-   Add authentication (JWT + OAuth providers)
-   Add deployment packaging (Docker, docker-compose) and basic CI checks
