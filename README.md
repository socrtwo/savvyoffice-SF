<!--MODERNIZED:v1-->
# Savvyoffice

> Migrated from SourceForge via SF2GH Migrator

[![Live page](https://img.shields.io/badge/live-page-ff2e93?style=for-the-badge)](https://socrtwo.github.io/savvyoffice-SF/)
[![Releases](https://img.shields.io/github/v/release/socrtwo/savvyoffice-SF?style=for-the-badge&color=7c3aed)](https://github.com/socrtwo/savvyoffice-SF/releases)
[![License](https://img.shields.io/github/license/socrtwo/savvyoffice-SF?style=for-the-badge&color=22d3ee)](https://github.com/socrtwo/savvyoffice-SF/blob/main/LICENSE)
[![Last commit](https://img.shields.io/github/last-commit/socrtwo/savvyoffice-SF?style=for-the-badge&color=34d399)](https://github.com/socrtwo/savvyoffice-SF/commits)

🌐 **Live:** https://socrtwo.github.io/savvyoffice-SF/  
📦 **Downloads:** [Releases](https://github.com/socrtwo/savvyoffice-SF/releases)  
📂 **Source:** [socrtwo/savvyoffice-SF](https://github.com/socrtwo/savvyoffice-SF)

---

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

---

## 📜 SourceForge heritage

This project originated on **SourceForge** before being migrated to GitHub. The legacy SourceForge entry, if still available, can be searched at:

🔗 https://sourceforge.net/projects/savvyoffice/

The repository here at `socrtwo/savvyoffice-SF` is the canonical, actively-maintained home. All future updates, issue tracking, and releases happen on GitHub.

## 🛠️ Contributing

Issues and pull requests are welcome at [https://github.com/socrtwo/savvyoffice-SF/issues](https://github.com/socrtwo/savvyoffice-SF/issues).

## 📝 License

See the [LICENSE](https://github.com/socrtwo/savvyoffice-SF/blob/main/LICENSE) file in this repository. If no license file is present, the project is shared as-is for reference and personal use; please contact the maintainer for other use cases.

---

*Maintained by [@socrtwo](https://github.com/socrtwo)*
