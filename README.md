# attackexcel

A command-line tool to work with MITRE ATT&CK and Excel

Attackexcel was born out of a necessity to create MITRE ATT&CK Heatmaps quickly and easily that
represent a variety of things. It does this through the use of two sub-commands: *seed*, and *layer*.
The former downloads the knowledgebase from the MITRE TAXII server into an Excel workbook and adds some
useful columns, while the latter transforms the workbook into a json file suitable for use
with the ATT&CK Navigator.

This tool works with version 9 of MITRE ATT&CK, version 4.3 of the navigator, and 4.2 of the layer files.
## Sub-command: seed
### Basic usage
`attackexcel seed <filename>`

This will download the Enterprise matrix into the named file, including all techniques and subtechniques
for every platform. It will create the following individual worksheets in an attempt to normalize the data.
  * techniques - including the technique ID, name, description, whether it is a subtechnique or not, and what tactics
    and platforms are associated with it.
  * dataSources - a unique list of the dataSources associated with at least one of the identified techniques.
  * techniquesToDataSources - Since the relationship between techniques and data sources is many-to-many, this is the
    necessary join table.

### Working with a different matrix
By default *seed* will download the Enterprise matrix. You can choose to download one of the other two matrices
by using the --domain switch.

This will download the ICS matrix:

`attackexcel seed [filename] --domain ics-attack`

This will download the Mobile matrix:

`attackexcel seed [filename] --domain mobile-attack`


### Filtering
By default, *seed* will include techniques associated with all platforms (even PRE!), but you can change this
behavior with the platforminclude or platformexclude switches and a space-separated list of platforms. If using
platforminclude, a given technique will be included if at least one of its associated platforms is in the list. If
using platformexclude, a given technique will be excluded only if *none* of its associated platforms are in the list.
These switches are mutually exclusive. Filter values are case-sensitive, and must be one of 'SaaS', 'macOS', 'PRE',
'IaaS', 'Linux', 'Office 365', 'Containers', 'Google Workspace', 'Windows', 'Network', 'Azure AD', 'Android', 'iOS',
'Field Controller/RTU/PLC/IED', 'Safety Instrumented System/Protection Relay', 'Control Server', 'Input/Output Server',
'Human-Machine Interface', 'Engineering Workstation', or 'Data Historian'. Values must be enclosed in quotes if it
contains a space. The supplied values must all be relevant to the chosen domain.

This will download techniques associated with Windows, macOS, Linux, and Office 365 from the Enterprise matrix:

`attackexcel seed [filename] --platforminclude Windows macOS Linux 'Office 365'`

This will download techniques associated with all platforms except PRE from the Enterprise matrix:

`attackexcel seed [filename] --platformexclude PRE`

This will download techniques associated with Control Server from the ICS matrix:

`attackexcel seed [filename] --domain ics-attack --platforminclude 'Control Server'`

This will download techniques associated with Android from the Mobile matrix:

`attackexcel seed [filename] --domain mobile-attack --platforminclude 'Android'`


### Subtechniques
By default, *seed* will include all subtechniques. You can change this behavior with the no-subtechniques flag.

`attackexcel seed [filename] --no-subtechniques`

## Sub-command: layer
### Basic usage
`attackexcel layer [input filename] [output filename]`

This will open the named input file, look for a worksheet named techniques, look for specific headers (see below), 
iterate through all rows except the first row, and use the values in each row to generate a json output file suitable
for use as a layer with the ATT&CK Navigator.

This tool looks for the following named column headers: techniqueID, color, enabled, score, comment. Only the values
in those columns will be included. Technically none of them are required, but the resulting json file will not work
with the Navigator unless at least techniqueID exists, and it won't be terribly useful if at least one of the others
is also present. If you are using this tool, then chances are you want to include at least *score*.

### Specifying a different worksheet
If, for whatever reason, the worksheet you want to use is not named *techniques*, you can specify which sheet to use
by using the worksheet switch.

`attackexcel layer [input filename] [output filename] --worksheet "ATT&CK Techniques"`

### Adding metadata
The Navigator supports a limited amount of metadata: name, and description. These can be added to the layer file
by using switches of the same name.

`attackexcel layer [input filename] [output filename] --name "My Layer" --description "My Description"`

### Specifying platforms
If you want the Navigator to only show a certain list of platforms, you can specify that at the command line as well
using the same switches as *seed*. Note, however, that all the techniques will still be in the layer file and
will be available within the Navigator. They will merely be hidden.

`attackexcel layer [input filename] [output filename] --platforminclude Windows MacOS Linux IaaS`
`attackexcel layer [input filename] [output filename] --platformexclude PRE`

## Sample worksheet
Seeding a workbook is cool, but if you really want to leverage the power of Excel and ATT&CK you'll want to manipulate
the data (most likely score, comment, and enabled) in a variety of ways. There is a sample worksheet
*enterprise-matrix.xlsx* in the GitHub repo that has some more advanced examples of this.