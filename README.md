# Savvy Repair for Microsoft Office

Repairs corrupt DOCX, XLSX, and PPTX files using 4 algorithmic methods: zip repair, strict XML validation truncation, lax validation, and text salvage.

## Screenshots

Visit the [SourceForge project page](https://sourceforge.net/projects/savvyoffice/) to view screenshots.

> **Tip:** If you have screenshots to contribute, open a PR adding them to a `screenshots/` folder!

**Language:** Delphi  
**License:** MIT

## Features

- Zip archive structure repair
- Strict XML validation with truncation
- Lax XML validation (recovers more data at the cost of some formatting)
- Plain text salvage as a last resort
- Works with all Office 2007+ formats

## System Requirements

- Windows XP or later
- Delphi 7 (for original build) or Free Pascal / Lazarus (free alternative)

## Installation & Usage

### Building from Source (Delphi 7)

1. Open the `.dpr` project file in Delphi 7
2. Press **F9** to compile and run

### Building with Free Pascal (free alternative)

```bash
sudo apt-get install fpc    # Linux
# or download from https://www.freepascal.org/
fpc -Sd src/*.pas
```

### Using a Pre-built Release

Download the latest release from the [Releases](../../releases) page.

## Origin

This project was originally hosted on SourceForge and has been migrated to GitHub for easier access and collaboration.

- **SourceForge:** [savvyoffice](https://sourceforge.net/projects/savvyoffice/)
- **Migrated with:** [SF2GH Migrator](https://github.com/socrtwo/sf-to-github)

## Contributing

Contributions are welcome! Feel free to:

1. Fork this repository
2. Create a feature branch (`git checkout -b my-feature`)
3. Commit your changes (`git commit -m "Add my feature"`)
4. Push to the branch (`git push origin my-feature`)
5. Open a Pull Request

## License

MIT License — see [LICENSE](LICENSE) for details.
