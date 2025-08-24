# Lean Digital Twin (Pro Edition)

A comprehensive Python application for managing Lean Digital Twin data with Excel integration, SPARQL querying, and graphical visualization.

## Features

- **4-Tab Interface**: Graphical Model, Asset Hierarchy, Functionalities, and Datasheet Editor
- **Excel Integration**: Embedded Excel windows with drag-and-drop functionality
- **SPARQL Integration**: Query semantic networks and repositories
- **Project Management**: Save/load projects as ZIP files
- **Theme Support**: Light/dark theme switching
- **Asset Hierarchy**: Tree-based management of Plant → Unit → Area → Equipment → Sub-Equipment → Asset

## Requirements

- **Windows OS** (required for Excel embedding)
- **Python 3.7+**
- **Microsoft Excel**
- **Required Python packages** (see requirements.txt)

## Installation

1. **Clone or download** the application files
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the application**:
   ```bash
   python lean_digital_twin.py
   ```

## Usage

### Getting Started

1. **Launch the application** - You'll see a welcome dialog
2. **Create a new project** or **open an existing one**
3. **Configure SPARQL endpoint** in the Graphical Model tab
4. **Fetch nodes** from your semantic network
5. **Build asset hierarchy** in the Asset Hierarchy tab
6. **Import Excel datasheets** in the Functionalities tab
7. **Map data** using the Datasheet Editor tab

### Key Workflows

#### 1. Setting up SPARQL Connection
- Go to **Graphical Model** tab
- Enter your **Semantic Network URL** (SPARQL endpoint)
- Set **SPARQL Prefix** (e.g., `PREFIX ex: <http://example.org/pumps#>`)
- Click **Test Connection**
- Click **Fetch All Nodes**

#### 2. Building Asset Hierarchy
- Go to **Asset Hierarchy** tab
- Right-click to add entities (Plant, Unit, Area, Equipment, etc.)
- Use context menus to create hierarchy
- Entities can be created manually or from repository nodes

#### 3. Managing Datasheets
- Go to **Functionalities** tab
- Import Excel templates using **Import Template**
- Create tags and associate with nodes
- Link datasheets to tags

#### 4. Data Mapping
- Go to **Datasheet Editor** tab
- Select a tag with associated datasheets
- Double-click properties to preview values
- Use drag-and-drop to copy values to Excel cells

### Keyboard Shortcuts

- `Ctrl+O`: Open Project
- `Ctrl+S`: Save Project
- `F5`: Update Data Model
- `Ctrl+F`: Focus filter (in applicable tabs)

## File Structure

```
app_directory/
├── lean_digital_twin.py    # Main application
├── requirements.txt        # Python dependencies
├── excel_files/           # Excel datasheets (created automatically)
├── logs/                  # Application logs (created automatically)
├── settings.json          # Application settings (created automatically)
└── tags.json              # Tag associations (created automatically)
```

## Configuration

### Settings (settings.json)
- `theme`: "light" or "dark"
- `repo_url`: SPARQL endpoint URL
- `sparql_prefix`: SPARQL namespace prefix
- `recent_repos`: List of recently used repositories

### Tags (tags.json)
- Maps tag names to associated nodes and datasheets
- Includes cell mapping information for drag-and-drop

## Troubleshooting

### Common Issues

1. **Excel not found**
   - Ensure Microsoft Excel is installed
   - Check that xlwings is properly installed

2. **SPARQL connection fails**
   - Verify the endpoint URL is correct
   - Check network connectivity
   - Ensure the SPARQL prefix matches your data

3. **Application crashes on startup**
   - Check Python version (3.7+ required)
   - Verify all dependencies are installed
   - Check Windows compatibility

### Logs

Application logs are stored in the `logs/` directory. Check `app.log` for detailed error information.

## Development

### Project Structure
The application is currently in a single file. Consider modularizing for better maintainability:

```
src/
├── ui/                    # UI components
├── core/                  # Core functionality
├── excel/                 # Excel integration
├── sparql/                # SPARQL integration
└── utils/                 # Utilities
```

### Testing
Add unit tests for core functionality:
```bash
pip install pytest
pytest tests/
```

## License

This application is designed for Lean Digital Twin management. Please ensure compliance with your organization's data handling policies.

## Support

For issues and questions:
1. Check the logs in `logs/app.log`
2. Verify all requirements are met
3. Test with sample data first
4. Ensure Windows compatibility