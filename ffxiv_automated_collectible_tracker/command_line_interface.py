import asyncio
import click
import logging
import yaml

from ffxiv_automated_collectible_tracker.ffxiv_gsheet_updater import update_spreadsheets
from ffxiv_automated_collectible_tracker.lodestone import get_fc_members_formatted_with_world


@click.group()
def cli():
    pass


@cli.command()
@click.argument("world", type=str)
@click.argument("fc-name", type=str)
@click.option("--characters-file", type=click.File("w"), default="characters.yaml")
def get_fc_members_list(world, fc_name, characters_file):
    """Create the yaml file of FC members."""
    loop = asyncio.get_event_loop()
    members = loop.run_until_complete(
        get_fc_members_formatted_with_world(world, fc_name)
    )
    yaml.safe_dump(members, characters_file)


@cli.command()
@click.option("--credentials-file", type=click.Path(exists=True), default=".credentials.json")
@click.option("--sheet-config-file", type=click.File("r"), default="config.yaml")
@click.option("--characters-file", type=click.File("r"), default="characters.yaml")
def fill_sheet_data(credentials_file, sheet_config_file, characters_file):
    """Fill in the character data on the Google Sheets."""
    sheets_config = yaml.safe_load(sheet_config_file)
    characters = yaml.safe_load(characters_file)
    loop = asyncio.get_event_loop()
    loop.run_until_complete(
        update_spreadsheets(credentials_file, sheets_config, characters)
    )


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(message)s', datefmt='%H:%M')
    cli()
