# HTML to PPTX Converter

A web-based application that converts HTML files to PowerPoint presentations (PPTX).

## Features

- Upload HTML files and convert them to PowerPoint presentations
- Simple web interface built with Flask
- Session management for user operations
- File handling with organized directories for uploads and outputs

## Project Structure

```
html-to-pptx/
├── app.py                 # Main Flask application file
├── requirements.txt       # Python dependencies
├── static/                # Static files (CSS, JS, images)
├── templates/            # HTML templates
├── uploads/              # Temporary storage for uploaded HTML files
├── output/               # Generated PowerPoint presentations
└── flask_session/        # Session data storage
```

## Installation

1. Clone or download this repository
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the Flask application:
   ```
   python app.py
   ```
2. Open your web browser and navigate to the provided URL (typically http://127.0.0.1:5000)
3. Upload your HTML file using the web interface
4. Convert and download the resulting PowerPoint presentation

## Dependencies

- Flask
- python-pptx
- Other dependencies as listed in requirements.txt

## License

This project is open source and available under the MIT License.
