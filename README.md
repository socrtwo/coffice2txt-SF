# Corrupt Office File Salvager

Extracts readable text from corrupt Microsoft Office files (DOC, XLS, PPT) when the application itself cannot open them. Uses low-level binary parsing to salvage whatever text remains.

**Language:** Perl  
**License:** MIT

## Features

- Extracts text from corrupt DOC, XLS, and PPT files
- Low-level binary parsing bypasses Office file format errors
- Works when Microsoft Office refuses to open the file
- Command-line interface for batch processing
- Outputs plain text for easy recovery

## System Requirements

- Perl 5.10 or later
- Linux, macOS, or Windows (with Strawberry Perl or WSL)

## Installation & Usage

### Running

```bash
# Install Perl (if not already installed)
# Linux/macOS: usually pre-installed
# Windows: download Strawberry Perl from https://strawberryperl.com/

# Run the script
perl <script_name>.pl [arguments]
```

### Dependencies

If the script uses CPAN modules, install them with:
```bash
cpan install Module::Name
```

## Origin

This project was originally hosted on SourceForge and has been migrated to GitHub for easier access and collaboration.

- **SourceForge:** [coffice2txt](https://sourceforge.net/projects/coffice2txt/)
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
