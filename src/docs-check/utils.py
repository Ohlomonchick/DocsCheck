from dataclasses import dataclass
from typing import List


@dataclass
class Message:
    text: str
    position: str
    standard: str


class Verdict:
    """Class for storing result"""
    ok: bool
    messages: List[Message]
    standard: str
    position: str

    def __init__(self, ok: bool = True, messages: List[Message] = None, position: str = None, standard: str = None):
        self.ok = ok
        self.messages = messages
        self.position = position
        self.standard = standard

        if messages is None:
            self.messages = []
        if position is None:
            self.position = ""
        if standard is None:
            self.standard = ""

    def add_message(self, message: str):
        self.messages.append(Message(message, position=self.position, standard=self.standard))
        self.ok = False

    def __add__(self, other):
        if self.position and not other.position:
            for i in range(len(other.messages)):
                other.messages[i].position = self.position

        self.messages += other.messages
        if not other.ok:
            self.ok = False
        return self
