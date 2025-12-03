# Convert Junos XML Configuration to Excel

This tool converts a Junos XML configuration file into an Excel workbook. It extracts addresses, address-sets, applications, application-sets, and security policies, placing each into its own sheet for easier review and documentation.

### Getting the XML configuration (optional)

If you do not already have an XML configuration file, you can export one from a Junos device. Connect to the device using SSH, enter the CLI, and run:

```
show configuration | display xml | no-more
```

This prints the full configuration in XML format. Copy everything between `<configuration>` and `</configuration>`, then save it as a `.xml` file.

### Installing requirements

Inside the `convert.xml.to.excel` directory, install the required Python modules:

```
pip install pandas xlsxwriter
```

or

```
python3 -m pip install pandas xlsxwriter
```

### Running the converter

Place your XML file in the same directory as `main.py`. A simple example:

```
convert.xml.to.excel
│   main.py
│   junos-fw1.xml
```

Run the script:

```
python main.py
```

The script automatically processes all `.xml` files in the folder and creates matching `.xlsx` files. For example:

```
junos-fw1.xml → junos-fw1.xlsx
```

### Output details

The resulting Excel file may include the following sheets depending on the content of the XML:

- Addresses  
- Address-sets  
- Policies  
- Applications  
- Application-sets  

Each sheet organizes the configuration data in a structured manner to make analysis easier.