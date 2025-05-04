## Screenshots

Here are some screenshots of DocDynamo in action:

### PDF Merger
![PDF Merger WebApp Interface](./static/assets/screenshots/pdf-merger.png)


Note: To add these screenshots:
1. Create a `screenshots` folder in `static/assets/`
2. Save your application screenshots as PNG files
3. Name them appropriately as shown in the image paths above
4. Place them in the screenshots folder


# DocDynamo - PDF Tools

A powerful, free, and open-source web application for handling PDF files with ease. Built with Flask and modern web technologies.

## Features

- **PDF Merging**: Combine multiple PDF files into a single document
- **PDF Compression**: Reduce PDF file size while maintaining quality
- **PPT to PDF Conversion**: Convert PowerPoint presentations to PDF format

## Tech Stack

- **Backend**: Python Flask
- **Frontend**: HTML, TailwindCSS, JavaScript
- **PDF Processing**: PyMuPDF (fitz)
- **PowerPoint Conversion**: Microsoft Office COM Automation

## Requirements

- Python >= 3.13
- Microsoft PowerPoint (for PPT conversion)
- Dependencies listed in pyproject.toml:
  - Flask
  - PyMuPDF
  - pywin32

## Installation

1. Clone the repository
```bash
git clone https://github.com/Pratham-Prog861/pdf-merger2.git
cd pdf-merger2
```

2. Create and activate a virtual environment
```bash
python -m venv venv
source venv/bin/activate  # On Windows, use 'venv\Scripts\activate'
```

3. Install the required dependencies
```bash
pip install -r requirements.txt
```

4. Run the application
```bash
python app.py
```

## Usage

1. Start the Flask Server
```bash
python app.py
```

2. Open your web browser and go to 
```bash
http://localhost:5000
```

3. Use the web interface to:

- Merge multiple PDF files
- Compress PDF files with adjustable compression levels
- Convert PowerPoint presentations to PDF

## Features in Detail

### PDF Merger
- Upload multiple PDF files
- Maintains original quality
- Automatic cleanup of temporary files
- Downloads merged file immediately

### PDF Compressor
- Adjustable compression levels (1-100)
- Smart image compression
- Text optimization
- Maintains readability while reducing file size

### PPT to PDF Converter
- Supports .ppt and .pptx files
- High-quality conversion
- Maintains formatting and layouts
- Automatic temporary file cleanup

## Development
The project structure is organized as follows:

pdf-merger2/
├── app.py              # Main Flask application
├── templates/          # HTML templates
│   └── index.html     # Main interface
├── static/            # Static assets
│   ├── css/          # Stylesheets
│   ├── js/           # JavaScript files
│   └── assets/       # Images and other assets
└── pyproject.toml    # Project dependencies

## Contributing
Contributions are welcome! Feel free to:

- Report bugs
- Suggest new features
- Submit pull requests

## License
This project is open source and available under the MIT License.

## Author
Developed by Pratham Darji

- GitHub: Pratham-Prog861
- LinkedIn: Pratham Darji

## Acknowledgments
- Built with modern web technologies
- Uses TailwindCSS for styling
- FontAwesome for icons