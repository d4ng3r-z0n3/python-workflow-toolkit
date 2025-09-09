# python-workflow-toolkit
Cross-platform Python utilities: universal dependency installer and screenshot-to-DOCX generator for streamlined development and documentation workflows.
**README.md** (Universal Python Installer)

# Universal Python Installer

An intelligent, cross-platform Python environment manager that automatically handles dependencies, virtual environments, and system-level requirements for Python applications.

## Overview

This installer eliminates the complexity of Python project setup by automatically detecting dependencies, creating isolated environments, and handling system-level package requirements across Windows, Linux, and macOS platforms.

## Features

### Automatic Dependency Detection
- Scans Python files using AST parsing to identify required packages
- Extracts system installation commands from script comments
- Maps common import names to correct PyPI package names
- Handles both standard library and third-party modules

### Cross-Platform Support
- **Windows**: Native Python detection with registry fallback
- **Linux**: Package manager integration (apt, dnf, pacman)
- **macOS**: Homebrew and system Python support
- Universal fallback mechanisms for all platforms

### Virtual Environment Management
- Creates isolated Python environments automatically
- Handles Python executable detection across different installations
- Configures proper PATH and environment variables
- Manages pip installations with version conflict resolution

### Enhanced Compatibility
- Fixes Tkinter directory issues on Windows
- Resolves X11 authorization problems on Linux
- Handles broken package dependencies automatically
- Provides multiple screenshot tool options for GUI applications

### Intelligent Error Handling
- Comprehensive debugging output with color-coded logging
- Automatic fallback to alternative installation methods
- Graceful degradation when system packages unavailable
- Detailed error reporting with suggested solutions

## Installation

No installation required. Simply place the `install` script in your project directory.

### System Requirements

- Python 3.6 or higher
- Administrative privileges for system package installation (Linux/macOS)
- Internet connection for package downloads

### Platform-Specific Dependencies

**Linux (Ubuntu/Debian):**
```bash
sudo apt update
sudo apt install python3-venv python3-pip
```

**Linux (CentOS/RHEL):**
```bash
sudo dnf install python3-venv python3-pip
```

**Windows:**
- Python installer from python.org includes all required components

**macOS:**
```bash
xcode-select --install  # For development tools
```

## Usage

### Basic Usage
```bash
python install [script_name.py] [script_arguments...]
```

### Examples
```bash
# Auto-detect Python script in current directory
python install

# Run specific script with arguments
python install my_app.py --config settings.json

# Run with debugging output
python install my_script.py --verbose
```

### Workflow

1. **Script Detection**: Automatically finds Python scripts or uses specified file
2. **Dependency Analysis**: Extracts imports and system requirements
3. **Environment Creation**: Sets up isolated virtual environment
4. **Package Installation**: Installs required dependencies with conflict resolution
5. **Script Execution**: Runs target script with proper environment configuration

## Advanced Features

### System Package Integration
The installer can execute system-level installation commands found in script comments:

```python
# sudo apt install imagemagick scrot
# sudo dnf install ImageMagick scrot
import subprocess
```

### Custom Package Mapping
Handles complex import-to-package mappings:
- `cv2` → `opencv-python`
- `PIL` → `Pillow`
- `sklearn` → `scikit-learn`
- And many more common cases

### Desktop Integration (Linux)
Automatically creates desktop launchers for GUI applications with proper environment configuration.

## Troubleshooting

### Common Issues

**Permission Errors:**
- Run with administrator privileges on Windows
- Use sudo for system package installation on Linux/macOS

**Virtual Environment Creation Fails:**
- Ensure Python venv module is installed
- Check available disk space
- Verify Python installation integrity

**Package Installation Errors:**
- Check internet connectivity
- Verify pip is up to date
- Review package availability on PyPI

**GUI Application Issues:**
- Install system display libraries (X11, Wayland)
- Configure proper DISPLAY environment variables
- Ensure graphics drivers are current

### Debug Mode
Enable verbose logging by setting environment variable:
```bash
export PYTHON_INSTALLER_DEBUG=1
python install my_script.py
```

## Security Considerations

This installer executes system commands found in Python script comments. Review all scripts before running with this installer, especially those from untrusted sources.

System commands are only executed after user confirmation, providing an additional security layer.

## License

Released under the GNU General Public License v3.0. Free to use, modify, and distribute according to GPL terms.

## Contributing

Contributions welcome through standard GitHub workflows. Please ensure:
- Cross-platform compatibility testing
- Comprehensive error handling
- Clear documentation for new features
- Backward compatibility maintenance

## Technical Details

### Architecture
- Single-file executable design for maximum portability
- Modular function architecture for easy maintenance
- Comprehensive logging system with color-coded output
- Robust error recovery mechanisms

### Supported Python Versions
- Python 3.6+ (primary support)
- Python 3.11+ (optimized performance)
- PyPy compatibility (experimental)

### Performance
- Parallel package installation when possible
- Cached virtual environment reuse
- Minimal overhead for subsequent runs
- Efficient dependency resolution algorithms
```

nsure all contributions maintain cross-platform compatibility and include appropriate documentation.
```
