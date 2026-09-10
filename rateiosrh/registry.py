"""Central registry for the PDF layout parsers available to the application."""

from typing import Dict, List, Type

from rateiosrh.core.exceptions import UnknownParserError
from rateiosrh.parsers.base import BaseParser
from rateiosrh.parsers.bradesco_dental import BradescoDentalParser
from rateiosrh.parsers.unimed import UnimedParser


PARSER_REGISTRY: Dict[str, Type[BaseParser]] = {
    UnimedParser.parser_id: UnimedParser,
    BradescoDentalParser.parser_id: BradescoDentalParser,
}


def get_parser(parser_id: str) -> BaseParser:
    """Return a fresh parser instance for a registered identifier."""

    try:
        parser_type = PARSER_REGISTRY[parser_id]
    except KeyError as exc:
        raise UnknownParserError(parser_id) from exc
    return parser_type()


def list_parsers() -> List[Dict[str, str]]:
    """List registered parsers in registry insertion order."""

    return [
        {"id": parser_id, "display_name": parser_type.display_name}
        for parser_id, parser_type in PARSER_REGISTRY.items()
    ]
