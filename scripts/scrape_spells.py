import os, pathlib, requests, time, re
from io import StringIO

import pandas as pd

from bs4 import BeautifulSoup

import logging
from customLogFormatter import CustomFormatter

log = logging.getLogger("scrape_spells.py")
log.setLevel(logging.INFO)

cons_handler = logging.StreamHandler()
cons_handler.setLevel(logging.INFO)
cons_handler.setFormatter(CustomFormatter())
log.addHandler(cons_handler)

ROOT_DIR = pathlib.Path(__file__).parent.parent.resolve()

CLASSES = ["Artificer", "Bard", "Cleric", "Druid", "Paladin", "Ranger", "Sorcerer", "Warlock", "Wizard"]
DND_WIKI = "https://dnd5e.wikidot.com"
DND_SPELLS_WIKI = "http://dnd5e.wikidot.com/spells"
MODIFIER_KEY = {
    "R": {
        "def": "Ritual",
        "loc": "Casting Time"
    },
    "D": {
        "def": "Dunamancy",
        "loc": "School"
    },    
    "DG": {
        "def": "Graviturgy Dunamancy",
        "loc": "School"
    },    
    "DC": {
        "def": "Chronurgy Dunamancy",
        "loc": "School"
    },    
    "HB": {
        "def": "Homebrew",
        "loc": "School"
    },
    "T": {
        "def": "Technomagic",
        "loc": "School"
    }
}


def scrape_spell_summary(output=os.path.join(ROOT_DIR,'output/queried/spell_table.csv')):
    r = requests.get(DND_SPELLS_WIKI)
    assert r.status_code == 200, f'DND_SPELLS_WIKI returned a bad request: {r.status_code}'

    html_text = r.text
    soup = BeautifulSoup(html_text, "html.parser")

    # collect all tables and concat them into one dataframe
    all_tables = soup.find_all('table')
    spell_table = None

    for i, t in enumerate(all_tables):
        df_t = pd.read_html(StringIO(t.prettify()))[0]

        hrefs = []
        for row in t.find_all('tr'):
            anchor = row.findChild('a')
            if anchor:
                # capture the link
                href = anchor.get('href')
                hrefs.append(DND_WIKI+href)

        df_with_links = df_t.assign(Level=[i]*df_t.shape[0], Links=hrefs)

        if spell_table is not None:
            spell_table = pd.concat([spell_table, df_with_links], ignore_index=True)
        else:
            spell_table = df_with_links
    
    # Write to CSV
    spell_table.to_csv(output, index=False)


def read_spell_csv(csv_path=os.path.join(ROOT_DIR,'output/queried/spell_table.csv')) -> pd.DataFrame:
    return pd.read_csv(csv_path, delimiter=',')


def scrape_spell_details(href: str):
    """
    
    :returns: [source_txt, classes_avail, material_component, description_text, tables, casting_time_val, range_val, duration_val]
    """

    r = requests.get(href, timeout=300)
    assert r.status_code == 200

    html_text = r.text
    soup = BeautifulSoup(html_text, 'html.parser')
    page_content = soup.find('div', {'id':'page-content'})

    # Outputs
    source_txt = None
    classes_avail = []
    description_text = []
    material_component = None
    tables = []

    all_children = page_content.findChildren(['p','ul','ol','table'])
    for i, child in enumerate(all_children):
        if i == 0:
            # SOURCE
            source_txt = child.text[len('Source:'):].strip()
        elif i == 1: 
            # SCHOOL + LEVEL
            # already known
            pass
        elif i == 2:
            # STATS, we just need the material component here
        
            # TODO: check against what we already have from the summary, or just copy it in
            is_casting_time = False
            is_range = False
            is_components = False
            is_duration = False

            casting_time_val, range_val, duration_val = '','',''

            for cont in child.contents:
                use_cont = cont.text
                if "casting time:" in use_cont.lower():
                    is_casting_time = True
                    continue
                if is_casting_time:
                    # handle
                    casting_time_val = use_cont.strip()
                    is_casting_time = False

                if "range:" in use_cont.lower():
                    is_range = True
                    continue
                if is_range:
                    # handle
                    range_val = use_cont.strip()
                    is_range = False

                if "components:" in use_cont.lower():
                    is_components = True
                    continue
                if is_components:
                    # handle
                    material_match = re.search('M (.*)', use_cont.strip())
                    if material_match:
                        material_component = material_match.group()[3:-1].strip()
                    is_components = False

                if "duration:" in use_cont.lower():
                    is_duration = True
                    continue
                if is_duration:
                    # handle
                    duration_val = use_cont.strip()
                    is_duration = False

            # idx_mat = child.text.find('M (')
            # if idx_mat > 0:
            #     # does exist
            #     idx_mat_end = child.text.find(')', idx_mat)
            #     material_component = child.text[idx_mat+3:idx_mat_end]

        elif i == len(all_children) - 1:
            # CLASSES
            classes = child.find_all('a')
            for c in classes:
                c_txt = c.text
                classes_avail.append((c_txt.split(' ')[0].strip(), 'optional' in c_txt.lower()))
        elif child.name == 'table':
            tables.append(child.prettify())
        else:
            # DESCRIPTION
            description_text.append(repr(child))

    # all_paragraphs = page_content.find_all('p')
    # for i, p in enumerate(all_paragraphs):
    #     if i == 0:
    #         # SOURCE
    #         source_txt = p.text[len('Source:'):].strip()
    #     elif i == 1: 
    #         # SCHOOL + LEVEL
    #         # already known
    #         pass
    #     elif i == 2:
    #         # STATS, we just need the material component here
    #         idx_mat = p.text.find('M (')
    #         if idx_mat > 0:
    #             # does exist
    #             idx_mat_end = p.text.find(')', idx_mat)
    #             material_component = p.text[idx_mat+3:idx_mat_end]
    #     elif i == len(all_paragraphs) - 1:
    #         # CLASSES
    #         classes = p.find_all('a')
    #         for c in classes:
    #             c_txt = c.text
    #             classes_avail.append((c_txt.split(' ')[0].strip(), 'optional' in c_txt.lower()))
    #     else:
    #         # DESCRIPTION
    #         description_text.append(repr(p))

    # all_tables = page_content.find_all('table')
    # for i, t in enumerate(all_tables):
    #     # capture the html of each table in the page content for later
    #     tables.append(t.prettify())

    return source_txt, classes_avail, material_component, description_text, tables, casting_time_val, range_val, duration_val

def scrape_all_spell_details(csv_path=os.path.join(ROOT_DIR,'output/queried/spell_table_detailed.csv'), 
                             output_path=os.path.join(ROOT_DIR,'output/queried/spell_table_detailed.csv')):
    df = read_spell_csv(csv_path)

    if "Queried" not in df.columns:
        # add the columns needed
        df["Source"] = ""
        for c in CLASSES: df[c] = ""
        df["Material Component"] = ""
        df["Description"] = ""
        df["Has Tables"] = False
        df["Queried"] = False
        df["Queried Casting Time"] = ""
        df["Queried Range"] = ""
        df["Queried Duration"] = ""

    table_count = 0
    table_errors = []

    for index, row in df.iterrows():
        spell_name = row["Spell Name"]
        log.info(f'[{index+1}/{df.shape[0]}] Getting spell "{spell_name}"')
        if row['Queried'] is True:
            log.info(f'Spell "{spell_name}" already successfully queried...skipping')
            continue

        href = row["Links"]
        try:
            source_txt, classes_avail, material_component, description_text, tables, casting_time_val, range_val, duration_val = scrape_spell_details(href)
        except BaseException as e:
            log.error(e)
            break

        df.loc[index, "Source"] = source_txt
        for c in CLASSES:
            for cls, opt in classes_avail:
                if c.lower()==cls.lower():
                    if opt:
                        df.loc[index,c] = "Optional"
                    else:
                        df.loc[index,c] = "Yes"

        if material_component is not None:
            df.loc[index, "Material Component"] = material_component
        if casting_time_val:
            df.loc[index, "Queried Casting Time"] = casting_time_val
            if casting_time_val.lower() != df.loc[index, "Casting Time"].lower().strip():
                log.debug(f"Detected Casting Time mismatch: {casting_time_val}")
        if range_val:
            df.loc[index, "Queried Range"] = range_val
            if range_val.lower() != df.loc[index, "Range"].lower().strip():
                log.debug(f"Detected Range mismatch: {range_val}")
        if duration_val:
            df.loc[index, "Queried Duration"] = duration_val
            if duration_val.lower() != df.loc[index, "Duration"].lower().strip():
                log.debug(f"Detected Duration mismatch: {duration_val}")

        # collect raw description paragraph separated by the pipe (|)
        df.loc[index, "Description"] = "|".join(description_text)

        # handle tables
        has_tables = bool(len(tables) > 0)
        if has_tables:
            log.info(f'Spell "{spell_name}" contains tables')
            table_count += 1

            for j, t in enumerate(tables):
                try:
                    with open(os.path.join(ROOT_DIR,f'./resources/tables/{spell_name}_table_{j}.html'), 'w', encoding='utf-8') as f:
                        f.write(t)
                except BaseException:
                    log.warning(f"Could not write table {j} for {spell_name}")
                    table_errors.append(f'{row["Spell_Name"]}_table_{j}')
        df.loc[index, "Has Tables"] = has_tables

        # yay, we queried it
        df.loc[index,"Queried"] = True

        # sleep for a tiny bit to not send too many requests in a short time
        sleep_for = 0.25
        log.debug(f'sleeping for {sleep_for}s')
        time.sleep(sleep_for)

    # write to csv as a backup
    df.to_csv(output_path, index=False)

    # give some details
    log.info(f'There were {table_count} tables in the total query')
    if table_errors:
        log.warning(f'Table errors: {table_errors}')

    return df

def move_superscripts_to_usable(df):
    """
    There are ritual and footnote-related details denoted as superscripts, place them into the dataframe.
    
    """
    
    # add Ritual and Notes columns (directly related to superscripts)
    df["Ritual"] = False
    df["Notes"] = ""

    r = requests.get(DND_SPELLS_WIKI)
    assert r.status_code == 200, f"DND_SPELLS_WIKI did not return a 200 status code. Returned: {r.status_code}"

    html_text = r.text
    soup = BeautifulSoup(html_text, "html.parser")
    soup.find_parent()

    # capture all superscripts
    all_sup = soup.find_all('sup')

    for sup in all_sup:
        # found a superscript, handle it
        superscript = sup.text
        row_parent = sup.find_parent('tr')

        # it wasn't in the table, it was probably from the key
        if row_parent is None: continue

        # find the right spell in the df
        spell_name = row_parent.findChild('td').text
        spell_index = df.index[df['Spell Name'] == spell_name].tolist()[0]

        # Details about the superscript key
        ref_dict = MODIFIER_KEY.get(superscript, None)
        if ref_dict is None:
            # unhandled but fine because all modifiers should have been defined.
            log.warning("AHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH")
            continue
        col_name = ref_dict['loc']

        # remove the superscript text from the df
        df.at[spell_index, col_name] = df.at[spell_index, col_name].replace(' '+superscript,'').strip()

        # add the information from the superscript to the correct column
        if superscript == "R":
            df.at[spell_index, "Ritual"] = True
        else:
            df.at[spell_index, "Notes"] = MODIFIER_KEY[superscript]["def"]
    return df

def split_out_components_and_conc(df):
    # split components and concentration

    # first make the columns
    df["Concentration"] = False
    df["Verbal"] = False
    df["Somatic"] = False
    df["Material"] = False

    for index, row in df.iterrows():
        if "V" in row["Components"]:
            df.loc[index,"Verbal"] = True
        if "S" in row["Components"]:
            df.loc[index,"Somatic"] = True
        if "M" in row["Components"]:
            df.loc[index,"Material"] = True
        if "Concentration" in row["Components"]:
            # remove Concentration from the Duration column and put it in its own column
            df.loc[index,"Duration"] = df.loc[index,"Duration"][len("Concentration"):].strip(',').strip()
            df.loc[index,"Concentration"] = True

    return df

def final_csv_export(df: pd.DataFrame, output_name: str) -> None:
    """
    Convert the working CSV DataFrame into a ready-to-use CSV file
    """
    # remove Components and Queried columns, we don't care about them now
    df.drop('Components',axis=1)
    df.drop('Queried',axis=1)

    # add in two new columns that create_cards.py cares about
    df['Generate Card'] = True
    df['Blurb'] = None

    # reorder columns (based on vibes)
    df = df[['Generate Card',
             'Spell Name',
             'School',
             'Casting Time',
             'Range',
             'Duration',
             'Ritual',
             'Concentration',
             'Verbal',
             'Somatic',
             'Material',
             'Level',
             'Artificer',
             'Bard',
             'Cleric',
             'Druid',
             'Paladin',
             'Ranger',
             'Sorcerer',
             'Warlock',
             'Wizard',
             'Material Component',
             'Blurb',
             'Description',
             'Has Tables',             
             'Links',
             'Source',
             'Notes',
             'Queried Casting Time',
             'Queried Range',
             'Queried Duration'
        ]]
    df.to_csv(output_name, index=False)

def convert_to_excel(input_csv, output_name):
    """
    Convert the working CSV file into a ready-to-use Excel (or ODS) file
    """
    df = pd.read_csv(input_csv)

    # remove Components and Queried columns, we don't care about them now
    df.drop('Components',axis=1)
    df.drop('Queried',axis=1)

    # add in two new columns that create_cards.py cares about
    df['Generate Card'] = True
    df['Blurb'] = None

    # reorder columns (based on vibes)
    df = df[['Generate Card',
             'Spell Name',
             'School',
             'Casting Time',
             'Range',
             'Duration',
             'Ritual',
             'Concentration',
             'Verbal',
             'Somatic',
             'Material',
             'Level',
             'Artificer',
             'Bard',
             'Cleric',
             'Druid',
             'Paladin',
             'Ranger',
             'Sorcerer',
             'Warlock',
             'Wizard',
             'Material Component',
             'Blurb',
             'Description',
             'Has Tables',             
             'Links',
             'Source',
             'Notes',
             'Queried Casting Time',
             'Queried Range',
             'Queried Duration'
        ]]

    # write it to an excel file
    df.to_excel(output_name, index=False)


def do_all_the_queries(final_output_file):
    """
    Perform all of the querying needed *from scratch* to produce the input Excel file create_cards.py expects.
    """
    # Make a directory to save intermediate products
    output_queried_dir = os.path.join(ROOT_DIR, 'output', 'queried')
    os.makedirs(output_queried_dir, exist_ok=True)

    log.info(f"Querying spell summaries from {DND_SPELLS_WIKI}")
    scrape_spell_summary(os.path.join(output_queried_dir,'spell_summary.csv'))

    log.info(f"Querying all spell details")
    df = scrape_all_spell_details(os.path.join(output_queried_dir,'spell_summary.csv'), os.path.join(output_queried_dir,'spell_table_detailed.csv'))
    
    log.info(f"Correcting superscripts")    
    # df = pd.read_csv(os.path.join(output_queried_dir, 'spell_table_detailed.csv'))
    df = move_superscripts_to_usable(df)
    
    log.info(f"Splitting components and concentration")
    df = split_out_components_and_conc(df)
    
    log.info(f"Producing the final CSV file at: '{final_output_file}'")
    final_csv_export(df, final_output_file)

    # # also write as an ods file for open-source file reading
    # convert_to_excel(os.path.join(output_queried_dir,'spell_table_final.csv'), final_output_file.replace(".xlsx", ".ods"))

    # log.info(f"Producing the final Excel file at: '{final_output_file}'")
    # convert_to_excel(os.path.join(output_queried_dir,'spell_table_final.csv'), final_output_file)

    # # also write as an ods file for open-source file reading
    # convert_to_excel(os.path.join(output_queried_dir,'spell_table_final.csv'), final_output_file.replace(".xlsx", ".ods"))


if __name__ == "__main__":
    csv_output_file = os.path.join(ROOT_DIR,'spell_list_inputs.csv')

    user_input = input(f"This will send many GET requests to the wiki and overwrite excel file '{csv_output_file}' over the course of ~15 minutes.\nDid you mean to start this? (Y/N): ")
    if user_input.lower()[0] == 'y':
        do_all_the_queries(csv_output_file)
    else:
        log.info("Exiting.")