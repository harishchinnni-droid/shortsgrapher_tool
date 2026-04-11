# Stock Automation Tool

A Python-based automation tool for stock management and operations.

## Project Structure

```
Stock_Automation/
├── src/                    # Source code directory
│   ├── __init__.py        # Package initialization
│   ├── main.py            # Main application logic
│   └── utils.py           # Utility functions
├── tests/                 # Unit tests
│   ├── __init__.py
│   └── test_main.py       # Main module tests
├── config/                # Configuration files
│   └── config.yml         # Application configuration
├── main.py                # Entry point
├── setup.py               # Package setup configuration
├── requirements.txt       # Project dependencies
├── .gitignore             # Git ignore rules
└── README.md              # This file
```

## Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd Stock_Automation
   ```

2. **Create a virtual environment**
   ```bash
   python -m venv venv
   ```

3. **Activate the virtual environment**
   - On Windows:
     ```bash
     venv\Scripts\activate
     ```
   - On macOS/Linux:
     ```bash
     source venv/bin/activate
     ```

4. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

5. **Install the package in development mode**
   ```bash
   pip install -e .
   ```

## Usage

### Running the Application

```bash
python main.py
```

### Configuration

Edit `config/config.yml` to customize application settings:
- Logging levels
- Database connection
- API timeouts
- Debug mode

### Running Tests

```bash
python -m pytest tests/
```

With coverage report:
```bash
pytest tests/ --cov=src --cov-report=html
```

## Development

### Code Style

Use black for code formatting:
```bash
black src/ tests/
```

Check code quality with flake8:
```bash
flake8 src/ tests/
```

Run type checking:
```bash
mypy src/
```

### Development Dependencies

Install dev dependencies:
```bash
pip install -e ".[dev]"
```

## API Reference

### StockAutomationTool

Main class for stock automation operations.

```python
from src.main import StockAutomationTool

# Initialize with config
tool = StockAutomationTool(config_path="config/config.yml")

# Run automation
tool.run()

# Stop automation
tool.stop()
```

## Logging

The application uses Python's standard logging module. Configure logging in `config/config.yml`:

```yaml
logging:
  level: "INFO"
  format: "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
  file: "logs/app.log"
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Write tests for new functionality
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions, please create an issue in the repository.

## Version History

- **1.0.0** - Initial release
  - Basic project structure
  - Configuration management
  - Logging setup
  - Unit test framework
a simple tool to automate trading
