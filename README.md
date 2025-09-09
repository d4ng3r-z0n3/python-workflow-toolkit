**README.md** (Cross-Platform Screenshot to DOCX Generator)

# Cross-Platform Screenshot to DOCX Generator

A professional documentation tool that captures screenshots and generates formatted Microsoft Word documents, designed for technical assignment submissions and penetration testing reports.

## Background

This application was created to streamline the documentation process for cybersecurity education assignments. Instead of manually formatting screenshots in Microsoft Word, students can capture, organize, and generate professional DOCX documents with proper formatting automatically.

## Installation

### Quick Start (Recommended)

**Using the Universal Installer:**
```bash
# Clone the repository
git clone https://github.com/d4ng3r-z0n3/python-workflow-toolkit.git
cd python-workflow-toolkit

# Run with automatic dependency installation
python install Screenshot.Docx.py
```

The installer will automatically:
- Create a virtual environment
- Install all required Python packages
- Handle system dependencies
- Launch the application

**Manual Installation**

**Step 1: System Requirements**

**All Platforms:**
- Python 3.6 or higher
- 4GB RAM minimum
- 100MB free disk space

**Linux (Ubuntu/Debian):**
```bash
sudo apt update
sudo apt install scrot xclip python3-tk python3-pip python3-venv
```

**Alternative Linux Tools:**
```bash
# Choose one based on your desktop environment
sudo apt install gnome-screenshot   # GNOME
sudo apt install spectacle          # KDE
sudo apt install flameshot          # Independent
sudo apt install imagemagick        # Command-line
```

**Windows:**
```bash
# Ensure pip is up to date
python -m pip install --upgrade pip
```
- No additional system requirements for screenshot capture
- Uses native Windows API for enhanced capture

**macOS:**
```bash
# Install Xcode command line tools if needed
xcode-select --install
```
- Uses built-in screencapture utility
- No additional system requirements

**Step 2: Python Dependencies**

Create virtual environment and install packages:
```bash
# Create virtual environment
python -m venv screenshot_env

# Activate virtual environment
# Windows:
screenshot_env\Scripts\activate
# Linux/macOS:
source screenshot_env/bin/activate

# Install required packages
pip install pyautogui Pillow python-docx requests
```

**Note for Windows users:** The `tkinter` package comes pre-installed with Python on Windows, so it doesn't need to be installed separately via pip.

## Features

### Advanced Screenshot Capture

**Platform-Specific Optimizations:**
- **Windows**: Direct window capture using Windows API with DWM frame detection
- **Linux**: Interactive selection using scrot with multiple fallback tools
- **macOS**: Native screencapture utility integration
- **Universal**: pyautogui fallback for all platforms

**Capture Options:**
- Configurable capture delay (1-10 seconds)
- Window-specific targeting
- Full screen capture
- Import existing images

### Professional Document Generation

- **Automatic Formatting**: Professional DOCX layout with customizable margins
- **Header Generation**: Student information, course codes, and module numbers
- **Section Organization**: Named sections with optional notes
- **Image Optimization**: Consistent sizing and center alignment
- **Page Management**: Automatic page breaks between sections

### Project Management

- **Save/Load Projects**: Preserve work sessions with .ssp project files
- **Screenshot Organization**: Drag-and-drop reordering interface
- **Metadata Editing**: Modify section names and notes after capture
- **Preview System**: Real-time screenshot preview with zoom controls

### User Interface

- **Tabbed Navigation**: Separate tabs for capture, editing, and settings
- **Modern Styling**: Professional appearance with intuitive controls
- **Responsive Design**: Adapts to different screen sizes and resolutions
- **Progress Tracking**: Visual feedback during document generation

## Security and Transparency

### Update Mechanism

This application includes an automatic update system that connects to external servers for version checking and user registration.

**Data Transmitted:**
- System information (hostname, OS, Python version)
- User registration details (name, email)
- Version checking and update downloads
- Usage statistics and error reporting

**Server Endpoints:**
- Update checks: `https://update.xn--mdaa.com/api/check-update`
- User registration: `https://update.xn--mdaa.com/api/register`
- Bug reports: `https://update.xn--mdaa.com/api/bug-report`

### Removing Update Features

If you prefer to run the application completely offline or have security concerns:

**Method 1 - Disable in Settings:**
1. Launch the application
2. Go to License menu
3. Uncheck "Auto Updates"
4. Settings are saved permanently

**Method 2 - Code Modification:**

Remove these sections from `Screenshot.Docx.py`:

```python
# Remove lines 38-179: UpdateChecker class
class UpdateChecker:
    # ... entire class definition

# Remove lines 334-339: updater initialization
self.updater = UpdateChecker(self)
if self.settings.get('auto_updates', True):
    self.root.after(2000, lambda: self.updater.check_for_updates(silent=True))

# Remove update menu items in create_menu() method
help_menu.add_command(label="Check for Updates", command=lambda: self.updater.check_for_updates(silent=False))
```

**Method 3 - Network Isolation:**
Block network access using firewall rules or run in an isolated environment.

The application functions completely without network connectivity once update features are disabled.

## Usage Guide

### Initial Setup

1. **Launch Application**
```bash
python install Screenshot.Docx.py
```

2. **Accept License Agreement**
   - Read and accept GNU GPL v3.0 terms
   - Provide registration information
   - Configure update preferences

3. **Configure Settings**
   - Enter personal information (name, course code)
   - Set default save location
   - Adjust document formatting preferences

### Capturing Screenshots

1. **Document Information**
   - Enter module/assignment number
   - Set document title
   - Configure capture delay

2. **Section Naming**
   - Pre-enter section name (optional)
   - Add descriptive notes
   - Use meaningful, descriptive names

3. **Capture Process**
   - Click "Capture Screenshot"
   - Follow platform-specific instructions
   - Confirm section name if not pre-entered

### Editing and Organization

1. **Screenshot Management**
   - View all captured screenshots in list
   - Reorder using up/down arrows
   - Delete unwanted captures
   - Edit section names and notes

2. **Preview System**
   - Select screenshots to preview
   - Zoom and scroll for detailed viewing
   - Edit metadata in real-time

### Document Generation

1. **Review Content**
   - Verify all screenshots and section names
   - Check notes and formatting preferences
   - Ensure proper order and organization

2. **Generate DOCX**
   - Click "Generate DOCX" button
   - Monitor progress bar
   - Choose to open file automatically

3. **File Management**
   - Documents saved with timestamp
   - Format: `FirstName.LastName.ModuleX_YYYYMMDD_HHMMSS.docx`
   - Saved to current working directory

## Professional Use Cases

### Educational Assignments
- Cybersecurity lab documentation
- Step-by-step process verification
- Proof of concept demonstrations
- Technical skill assessments

### Penetration Testing
- Vulnerability documentation
- Exploitation proof screenshots
- Report generation for clients
- Compliance documentation

### Technical Documentation
- Software installation guides
- Configuration procedures
- Troubleshooting documentation
- Training material creation

## Troubleshooting

### Installation Issues

**Installer fails:**
```bash
# Try with specific Python version
python3 install Screenshot.Docx.py

# Check Python installation
python --version
python3 --version

# Manual dependency installation
pip install --user tkinter pyautogui Pillow python-docx requests
```

**Permission errors:**
```bash
# Linux/macOS - run system commands manually
sudo apt install scrot xclip python3-tk

# Windows - run as administrator
# Right-click command prompt -> "Run as administrator"
```

### Screenshot Issues

**Linux - scrot not found:**
```bash
sudo apt install scrot
# Or use alternatives: gnome-screenshot, spectacle, flameshot
```

**Windows - Capture fails:**
- Run as administrator for system windows
- Check Windows security settings
- Verify graphics drivers are current

**All Platforms - General capture issues:**
- Increase capture delay
- Use import feature for existing images
- Check pyautogui fallback functionality

### Document Generation Problems

**Permission denied errors:**
- Close any open DOCX files with same name
- Check write permissions in target directory
- Run application with appropriate privileges

**Formatting issues:**
- Adjust image height in settings
- Modify page margins
- Check document template compatibility

**Memory errors with large images:**
- Reduce image dimensions before import
- Use lower capture resolution
- Process screenshots in smaller batches

## Development and Customization

### Code Structure
- Modular class-based architecture
- Platform-specific capture implementations
- Separation of GUI and document generation logic
- Comprehensive error handling throughout

### Customization Options
- Modify DOCX templates and styling
- Add new screenshot capture methods
- Extend metadata fields
- Customize GUI appearance and layout

### API Integration
- Document generation can be used programmatically
- Screenshot capture methods available independently
- Settings system supports external configuration

## License

Released under GNU General Public License v3.0. This ensures the software remains free and open source, allowing users to modify, redistribute, and improve the application according to their needs.

## Privacy and Data Handling

### Local Data Storage
- Screenshots stored locally during editing
- Project files contain only metadata and file references
- Settings saved in local JSON configuration files
- No permanent data transmission to external servers

### Optional Data Sharing
- Update checks can be disabled completely
- User registration information used only for support
- System information helps improve cross-platform compatibility
- All data transmission can be prevented through settings or code modification

## Contributing

Contributions welcome through standard GitHub workflows. Priority areas:
- Additional screenshot tools for various Linux distributions
- Enhanced DOCX formatting options
- Performance optimizations for large image sets
- Accessibility improvements for users with disabilities

Please ensure all contributions maintain cross-platform compatibility and include appropriate documentation.
