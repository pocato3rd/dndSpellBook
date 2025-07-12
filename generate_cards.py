import argparse, os, pathlib
import pandas as pd
import logging

from scripts.customLogFormatter import CustomFormatter
from scripts.create_cards import create_filtered_cards, Card

ROOT_DIR = pathlib.Path(__file__).parent.resolve()

log = logging.getLogger("generate_cards.py")
log.setLevel(logging.INFO)

cons_handler = logging.StreamHandler()
cons_handler.setLevel(logging.INFO)
cons_handler.setFormatter(CustomFormatter())
log.addHandler(cons_handler)


def main():
    parser = argparse.ArgumentParser(description = "A DND 5E spell card generator tool!")

    # defining arguments for parser object
    parser.add_argument("-p", "--preview", required=False, action='store_true',
            help=f"(Optional) Flag to indicate that you would like a preview of the filter. Will not create cards, just outputs number of cards that would be made"
    )
    parser.add_argument("-c", "--classes", type=str,
            metavar='class_list', default=None, required=False,
            help=f"(Optional) Comma-separated list of classes to filter on, overrides 'Generate Card' filter. ANDs with level list. Supported classes: {Card.CLASSES}"
    )
    parser.add_argument("-l", "--levels", type=str,
        metavar='level_list', default=None, required=False,
        help=f"(Optional) Comma-separated list of spell levels to filter on, overrides 'Generate Card' filter. ANDs with class list. Supported levels: 0 through 9, inclusive"
    )
    parser.add_argument("-i", "--input_file", type=str, required=False,
        metavar='input_file', default=os.path.join(ROOT_DIR, 'spell_list_inputs.xlsx'),
        help=f"(Optional) The .xlsx or .ods input file to pull spell details from. If --classes and --levels are not specified here, the 'Generate Cards' column will be used to filter on. Defaults to '{os.path.join(ROOT_DIR, 'spell_list_inputs.xlsx')}'"
    )
    parser.add_argument("-o", "--output_dir", type=str, required=False,
        metavar='output_dir', default=os.path.join(ROOT_DIR, 'output/cards'),
        help=f"(Optional) The output directory to put cards into. Cards will be further organized by level directories. If the output directory does not exist, it will be created recursively. Defaults to '{os.path.join(ROOT_DIR, 'output/cards')}'"
    )
    
    # parse the arguments from standard input
    args = parser.parse_args()

    df = pd.read_excel(args.input_file)
    apply_filters = []
     
    if args.classes != None:
        all_classes_filtered = None
        for c in args.classes.split(','):
            use_c = c.capitalize()
            if use_c not in Card.CLASSES: 
                log.error(f"Class '{use_c}' could not be parsed. Skipping. Available classes are:\n\t{Card.CLASSES}")
            else:
                class_filter = (df[use_c] == "Yes") | (df[use_c] == "Optional")
                
                if all_classes_filtered is None:
                    all_classes_filtered = class_filter
                else: 
                    all_classes_filtered = (all_classes_filtered) | class_filter
        if all_classes_filtered is not None: apply_filters.append(all_classes_filtered)

    if args.levels != None:
        all_levels_filtered = None
        for l in args.levels.split(','):
            use_l = int(l)
            if use_l > 9: 
                log.error(f"Level '{use_l}' could not be parsed. Levels must be 0-9, inclusive. Skipping")
            else:
                level_filter = (df['Level'] == use_l)
                
                if all_levels_filtered is None:
                    all_levels_filtered = level_filter
                else: 
                    all_levels_filtered = (all_levels_filtered) | level_filter
        if all_levels_filtered is not None: apply_filters.append(all_levels_filtered)

    if len(apply_filters) == 1:
        # just one filter, use it
        filtered_df = df[apply_filters[0]]
    elif len(apply_filters) == 2:
        # AND the two filters
        filtered_df = df[(apply_filters[0]) & (apply_filters[1])]
    else:
        # use 'Generate Card' column
        filtered_df = df[df['Generate Card']]
    
    if filtered_df.shape[0]:
        
        if args.preview:
            log.info(f"This query would create {filtered_df.shape[0]} cards. Preview mode enabled so exiting without creating the cards")
        else:
            log.info(f"Creating {filtered_df.shape[0]} cards...")
            create_filtered_cards(filtered_df, output_dir=args.output_dir)
    else:
        log.warning(f"0 cards selected with current filters. Exiting.")

if __name__ == "__main__":
    main()
