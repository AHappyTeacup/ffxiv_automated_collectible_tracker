import logging
from ffxiv_automated_collectible_tracker.command_line_interface import cli

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(message)s', datefmt='%H:%M')
    cli()
