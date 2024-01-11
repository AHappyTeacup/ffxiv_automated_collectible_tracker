"""ffxivapi is a community driven resource that I could have used.
In fact, I did when I started this project. But the achievement list for characters was incomplete, so I went direct to
source.
"""
import asyncio
import aiohttp
import copy
import logging
import re
import requests

from bs4 import BeautifulSoup as bs
from pathlib import PurePosixPath


logger = logging.getLogger(__name__)

_website_url = "eu.finalfantasyxiv.com"
_lodestone_url = f"{_website_url}/lodestone"
char_uri = f"https://{_lodestone_url}/character"
fc_url = f"https://{_lodestone_url}/freecompany"
achievement_name_regex = re.compile('^.*\sachievement\s"(?P<achievement_name>.*)"\searned!$')


async def get_char_details(char_name: str, world: str, achievements: bool = False, mounts: bool = False) -> dict:
    """Make queries for a characters ID, achievements, and mounts.

    :param char_name: The full name of the character.
    :param world: The name of the world that character is from.
    :param achievements: Boolean. Get the character's achievements.
    :param mounts: Boolean. Get the character's mounts.
    :return: Dictionary. All the requested data.
    """
    char_id = get_char_id(char_name, world)
    char_details = {"ID": char_id, "Name": char_name, "World": world}
    if achievements:
        char_achievements = await get_char_achievements(char_id)
        char_details["Achievements"] = char_achievements
    if mounts:
        char_mounts = await get_char_mounts(char_id)
        char_details["Mounts"] = char_mounts
    return char_details


def get_char_id(char_name: str, world: str) -> str:
    """Get the lodestone ID for a character.

    :param char_name: The full name of the character.
    :param world: The name of the world that character is from.
    :return: String. The lodestone ID for a character.
    """
    logger.info(f"Getting character id for '{char_name}'.")
    response = requests.get(char_uri, params={"q": f"\"{char_name}\"", "worldname": world})
    response_soup = bs(response.text, "html.parser")
    entries = response_soup.find_all("a", class_="entry__link")
    try:
        char_entry = [entry for entry in entries if entry.find("p", class_="entry__name").text == char_name][0]
        char_id = PurePosixPath(char_entry["href"]).parts[-1]
        return char_id
    except IndexError as e:
        logger.error(e)


async def _get_url_soup(session: aiohttp.ClientSession, url_list: [str]) -> [bs]:
    """Asynchronously pop a url from the queue, request it, and run BeautifulSoup on the results.

    :param session: An async http session.
    :param url_list: The list of URLs to be queried.
    :return: BeautifulSoup.
    """
    soups = []
    while url_list:
        url = url_list.pop()
        logger.info(f"Getting URL: {url}...")
        data = None
        while data is None:
            async with session.get(url) as response:
                data = await response.text()
            if response.status == 429:
                await asyncio.sleep(1)
                logger.info(f"Too Many Requests. Retrying {url}...")
                data = None
        soup = bs(data, "html.parser")
        soups.append(soup)
    return soups


async def _batch_get_url_soups(url_list: [str]) -> [bs]:
    """Asynchronously get a list of BeautifulSoups for a list of URLs.

    :param url_list: A list of URL strings.
    :return:  A list of BeautifulSoups.
    """
    this_url_list = copy.deepcopy(url_list)
    soups = []
    async with aiohttp.ClientSession() as session:
        soup_gathering_tasks = [_get_url_soup(session, this_url_list) for _ in range(10)]
        [soups.extend(subset_of_soups) for subset_of_soups in await asyncio.gather(*soup_gathering_tasks)]
    return soups


async def get_char_mounts(char_id: str) -> [str]:
    """Given a lodestone character ID, get the complete list of human readable names of their mount collection.

    :param char_id: The lodestone ID for a character.
    :return: A list of the human readable names of that character's mount collection.
    """
    logger.info(f"Getting mounts for '{char_id}'.")
    char_mounts_url = f"{char_uri}/{char_id}/mount"
    response = requests.get(char_mounts_url)
    response_soup = bs(response.text, "html.parser")
    mount_lis = response_soup.find_all("li", attrs={"data-tooltip_href": True}, class_="mount__list_icon")
    mount_urls = ["https://" + _website_url + mount_li.attrs["data-tooltip_href"] for mount_li in mount_lis]
    mounts = [soup.h4.text for soup in await _batch_get_url_soups(mount_urls)]
    return mounts


async def get_char_achievements(char_id: str) -> [str]:
    """Given a lodestone character ID, get the complete list of human readable names of their achievements.

    :param char_id: The lodestone ID for a character.
    :return: A list of the human readable names of that character's achievements.
    """
    logger.info(f"Getting achievements for '{char_id}'.")
    char_acvhievements_url = f"{char_uri}/{char_id}/achievement"
    response = requests.get(char_acvhievements_url)
    response_soup = bs(response.text, "html.parser")

    pages_li = response_soup.find("li", class_="btn__pager__current")
    if not pages_li:
        return []

    total_pages = int(pages_li.text.split()[-1])
    achievement_urls = [char_acvhievements_url + "/?page=%s" % (page_num + 1) for page_num in range(total_pages)]

    achievements = [
        achievement_name_regex.fullmatch(p.text).group("achievement_name")
        for soup in await _batch_get_url_soups(achievement_urls)
        for p in soup.find_all("p", class_="entry__activity__txt")
    ]

    return achievements


async def get_fc_members(world: str, fc_name: str) -> [str]:
    """Get the names of all of the members of a Free Company.

    :param world: The name of the world the FC is on.
    :param fc_name:The name of the FC
    :return: The list of names of the members of the FC.
    """
    logger.info(f"Getting Free Company '{fc_name}'.")
    response = requests.get(fc_url, params={"q": fc_name, "worldname": world})
    response_soup = bs(response.text, "html.parser")
    search_results = response_soup.find("div", class_="ldst__window")
    entries = search_results.find_all("a", class_="entry__block")
    char_entry = [entry for entry in entries if entry.find("p", class_="entry__name").text == fc_name][0]
    fc_id = PurePosixPath(char_entry["href"]).parts[-1]

    logger.info(f"Getting members for '{fc_id}'.")
    fc_members_url = f"{fc_url}/{fc_id}/member"

    response = requests.get(fc_members_url)
    response_soup = bs(response.text, "html.parser")

    response_soup.find("li", class_="btn__pager__current")
    total_pages = int(response_soup.find("li", class_="btn__pager__current").text.split()[-1])

    members_urls = [fc_members_url + "/?page=%s" % (page_num + 1) for page_num in range(total_pages)]

    members = [
        p.text
        for soup in await _batch_get_url_soups(members_urls)
        for p in soup.find("div", class_="ldst__window").find_all("p", class_="entry__name")
    ]

    return members


async def get_fc_members_formatted_with_world(world: str, fc_name: str) -> [str]:
    """Get the names of all of the members of a Free Company, in the form {name}@{world}.

    :param world: The name of the world the FC is on.
    :param fc_name:The name of the FC
    :return: The list of names of the members of the FC, in the form {name}@{world}.
    """
    return ["%s@%s" % (member, world) for member in await get_fc_members(world, fc_name)]
