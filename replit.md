# Overview

This is a Flask web application that processes Excel files (.xlsx and .xls formats) and extracts data for form filling purposes. The application provides a simple web interface where users can upload Excel files, and the system processes them using pandas to extract and display information about the file structure and contents. The application is designed with a Spanish language interface and uses Bootstrap for responsive UI components.

# User Preferences

Preferred communication style: Simple, everyday language.

# System Architecture

## Web Framework
- **Flask**: Chosen as the lightweight web framework for its simplicity and rapid development capabilities
- **Werkzeug middleware**: ProxyFix middleware configured for proper header handling in production environments
- **Session management**: Uses Flask's built-in session handling with configurable secret key

## Frontend Architecture
- **Template engine**: Jinja2 templating with base template inheritance for consistent UI
- **CSS framework**: Bootstrap 5 with dark theme for modern, responsive design
- **Icons**: Feather icons for consistent iconography
- **JavaScript**: Vanilla JavaScript for form validation, file upload handling, and UI interactions

## File Processing
- **Upload handling**: Secure file uploads with filename sanitization using Werkzeug
- **File validation**: Strict validation for file extensions (.xlsx, .xls) and size limits (16MB maximum)
- **Excel processing**: pandas library for reading and parsing Excel files
- **Data extraction**: Converts Excel data to Python dictionaries for easier form population

## Data Storage
- **File storage**: Local filesystem storage in 'uploads' directory for temporary file processing
- **No persistent database**: Application processes files on-demand without permanent storage

## Security Features
- **File upload security**: Secure filename handling and extension validation
- **Size limits**: Maximum file size restrictions to prevent abuse
- **Session security**: Configurable session secret key with environment variable support

## Configuration Management
- **Environment variables**: Uses environment variables for sensitive configuration (SESSION_SECRET)
- **Logging**: Built-in Python logging configured for debugging
- **Development vs Production**: Configurable settings for different deployment environments

# External Dependencies

## Python Libraries
- **Flask**: Web framework and routing
- **pandas**: Excel file reading and data manipulation
- **Werkzeug**: Secure file handling utilities
- **openpyxl/xlrd**: Excel file format support (via pandas)

## Frontend Resources
- **Bootstrap CSS**: CDN-hosted Bootstrap 5 with custom dark theme
- **Feather Icons**: CDN-hosted icon library
- **JavaScript**: Bootstrap JavaScript components for interactive elements

## Development Tools
- **Python logging**: Built-in logging for application monitoring
- **Flask development server**: Built-in development server for testing

## File System Dependencies
- **Local uploads directory**: Requires writable filesystem access for temporary file storage
- **Static file serving**: Flask's built-in static file serving for CSS/JS assets