# SpellBook

View markdown with: [markdownlivepreview.com](https://markdownlivepreview.com/)

This tool will produce `.docx` files for each given DND 5E spell. The data will be pulled from the provided `--input_file` (most likely `./spell_list_inputs.csv`) which was produced by scraping spells from [https://dnd5e.wikidot.com](https://dnd5e.wikidot.com).

The output cards will be saved by spell level in `--output_dir`, defaulting to `./output/cards`.

It is strongly recommended that Microsoft Word is used to open the `.docx` files for proper formatting (and printing). Exporting to `.png` is in work ([Issue #2](https://github.com/pocato3rd/dndSpellBook/issues/2) in this repository)

## Environment

This tool requires Python 3 installed with a set of required modules.

Python can be downloaded from [https://www.python.org/downloads](https://www.python.org/downloads/)

Once Python is installed, a [virtual environment](https://docs.python.org/3/library/venv.html) can be setup in the directory. Clone this repository then navigate to its directory in a terminal.

Then create a the virtual environment using:

```bash
    python -m venv ./venv
```

*A brief note on Windows versus Mac/Linux virtual environments:*

* The executable files should be the same between different operating systems but the intermediate folder will be different
   * Virtual environments generated on Windows will use `./venv/Scripts/`
   * Those created on Mac/Linux (Posix) will use `./venv/bin/`
* This will be called out explicitly for the requirement installation and usage but will only the Windows path will be used in the following examples


Once the virtual environment is created, modules can be installed by either 1. or 2. below:

1. Call pip directly:
   * WINDOWS: `./venv/Scripts/pip3.exe install -r requirements.txt`
   * POSIX: `./venv/bin/pip3.exe install -r requirements.txt`
2. Activate the venv, then install:
    * Activate:
        * WINDOWS: `./venv/Scripts/activate.bat`
        * POSIX: `source ./venv/bin/activate`
    * Install:
        * `pip3 install -r requirements.txt`

Then the tool can be run by calling Python directly:

```bash
    ./venv/Scripts/python.exe generate_cards.py [OPTIONS]

    ./venv/bin/python.exe generate_cards.py [OPTIONS]
```

## Usage

Run the script to create cards by:

```bash
    ./venv/Scripts/python.exe generate_cards.py [OPTIONS]
```

### Help usage:

```
usage: generate_cards.py [-h] [-p] [-c class_list] [-l level_list] [-i input_file] [-o output_dir]

A DND 5E spell card generator tool!

options:
  -h, --help            show this help message and exit
  -p, --preview         (Optional) Flag to indicate that you would like a preview 
                        of the filter. Will not create cards, just outputs number 
                        of cards that would be made.
  -c class_list, --classes class_list
                        (Optional) Comma-separated list of classes to filter on, 
                        overrides 'Generate Card' filter. ANDs with level list. 
                        Supported classes: ['Artificer', 'Bard', 'Cleric', 
                        'Druid', 'Paladin', 'Ranger', 'Sorcerer', 'Warlock', 'Wizard']
  -l level_list, --levels level_list
                        (Optional) Comma-separated list of spell levels to filter on, 
                        overrides 'Generate Card' filter. ANDs with class list. 
                        Supported levels: 0 through 9, inclusive
  -i input_file, --input_file input_file
                        (Optional) The .csv, .xlsx, .ods input file to pull spell 
                        details from. If --classes and --levels are not specified
                        here, the 'Generate Cards' column will be used to filter on. 
                        Defaults to './spell_list_inputs.csv'
  -o output_dir, --output_dir output_dir
                        (Optional) The output directory to put cards into. 
                        Cards will be further organized by level directories. 
                        If the output directory does not exist, it will be created recursively. 
                        Defaults to './output/cards'
```

### Examples: 

Generate all 0th, 1st, and 9th level spells:

```bash
./venv/Scripts/python.exe generate_cards.py --levels 0,1,9
```

Generate all 1st level spells that are either Paladin or Artificer:

```bash
./venv/Scripts/python.exe generate_cards.py --levels 1 --classes Artificer,Paladin
```

Generate spells based on the `=TRUE()` values of 'Generate Card' in `./spell_list_inputs.csv`. Place them in output directory `./output/my_character_cards`:

```bash
./venv/Scripts/python.exe generate_cards.py --input_file ./spell_list_inputs.csv --output_dir ./output/my_character_cards
```

## Warning

Do not manipulate anything in the `resources` folder. That folder is pretty load-bearing


