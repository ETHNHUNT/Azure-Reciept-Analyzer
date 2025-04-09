# Azure-Reciept-Analyzer
A Streamlit application using Azure Vision to analyze receipts.

# Project Title (Replace with actual title)

This project provides a tool for analyzing receipts using Azure Vision services.

## Setup

1.  **Clone the repository:**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```
    *(Replace `<repository_url>` and `<repository_directory>` with actual values if applicable)*

2.  **Install dependencies:**
    Ensure you have Python installed. Then, install the required packages:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Configure Environment Variables:**
    *   Open the `.env` file and fill in your Azure Vision Endpoint and API Key:
        ```env
        # Azure Vision Configuration
        VISION_ENDPOINT=<Your Azure Vision Endpoint>
        VISION_API_KEY=<Your Azure Vision API Key>

        # Processing Configuration (Defaults are likely fine)
        MAX_THREADS=3
        MAX_RETRIES=3
        POLLING_INTERVAL=1.5

        # Output Configuration (Defaults are likely fine)
        OUTPUT_DIR=./output
        LOG_LEVEL=INFO
        ```

## Running the Application

Once the setup is complete, you can run the Streamlit application:

```bash
streamlit run streamlit_app.py
```

This will start the application, and you can typically access it in your web browser at `http://localhost:8501`. 
