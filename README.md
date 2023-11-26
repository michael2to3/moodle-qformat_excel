# moodle-qformat_xlsx

`moodle-qformat_xlsx` is a plugin for Moodle that provides seamless integration for importing and exporting XLSX files. This plugin enhances the functionality of Moodle by allowing users to easily upload and download XLSX files, which can be particularly useful for managing course data, student information, grades, and more.

## Features

- **Import XLSX**: Easily import data from XLSX files into Moodle.
- **Export to XLSX**: Export Moodle data like grades, attendance, and course information into an XLSX file.
- **User-Friendly Interface**: Simple and intuitive UI for both importing and exporting processes.
- **Data Integrity**: Ensures that the data structure is maintained during the import/export process.

## Installation

1. **Download the Plugin**: Clone this repository or download the ZIP file.
    ```
    git clone https://github.com/michael2to3/moodle-qformat_xlsx.git
    ```
2. **Install the Plugin on Moodle**:
   - Navigate to your Moodle installation's `admin/tool/installaddon/index.php` page.
   - Upload the downloaded plugin ZIP file.
   - Follow the on-screen instructions to complete the installation.

3. **Verify Installation**:
   - After installation, 'moodle-qformat_xlsx' should be listed under the Plugins section in Moodle.

## Usage

### Importing XLSX Files

1. Navigate to the `moodle-qformat_xlsx` import section in Moodle.
2. Choose the XLSX file you want to import.
3. Map the columns in your XLSX file to the corresponding fields in Moodle.
4. Click on 'Import' to complete the process.

### Exporting to XLSX

1. Navigate to the section where you want to export data (e.g., grades, attendance).
2. Select the `Export to XLSX` option.
3. Customize the data fields and filters as per your requirements.
4. Click on 'Export' to download the XLSX file.

## Configuration

No additional configuration is needed after installation. However, you can modify the import/export settings based on your Moodle setup if necessary.

## Support

For issues, feature requests, or contributions, please use the [Issues](https://github.com/michael2to3/moodle-qformat_xlsx/issues) section of this repository.

## License

This plugin is released under the [GNU General Public License](LICENSE). See the LICENSE file for more details.
