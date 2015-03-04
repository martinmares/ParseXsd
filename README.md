# ParseXsd

Read input XSD (XMLSchema) file, parse it and write output to XLSX (Excel) or Console (stdout).

# $ parsexsd --help

```
parsexsd v0.1beta (c) 2015 Martin Mareš
Options:
  -x, --xsd=<s>                  name of the input XSD file
  -l, --xlsx=<s>                 name of the output XLSX file
  -s, --xlsx-enums=<s>           name of the output XLSX file for ENUMs
  -t, --stdout                   write the XSD structure on the screen
  -i, --indent                   name the elements in XLSX will be indented
  -b, --border                   generate a border for cells in XLSX?
  -c, --columns=<s>              the list of columns in the XLSX
  -m, --imports=<s>              on/off xsd:import tags (default: on) (default:
                                 on)
  -f, --frozen                   add "frozen" row and column started at A1
                                 position
  -r, --request-end-with=<s>     mark the line ending at {Request}
  -e, --response-end-with=<s>    mark the line ending at {Response}
  -h, --header-request           add a header to each of the {Request} elem.
  -a, --header-response          add a header to each of the {Response} elem.
  -u, --auto-filter              turn on the "auto filter on the first row"
  -o, --font-name=<s>            change the font (default: "Tahoma")
  -n, --font-size=<i>            change the font size (default 9)
  -d, --header-font-size=<i>     change the heder font size (default: 9)
  -v, --version                  Print version and exit
  -p, --help                     Show this message
```

# Sample terminal output with --stdout

## examples\msdn\print_stdout.sh

![msdn.xsd terminal output](https://raw.githubusercontent.com/martinmares/ParseXsd/master/images/msdn_xsd_stdout.png)

