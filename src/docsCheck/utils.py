from dataclasses import dataclass, field
from typing import List
from enum import Enum


class MessageTypes(Enum):
    ERROR = 0
    WARNING = 1


@dataclass
class Message:
    text: str
    position: str
    standard: str
    message_type: MessageTypes = field(repr=False)


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

    def add_message(self, message: str, position: str = None, message_type: MessageTypes = MessageTypes.ERROR):
        if position is None:
            self.messages.append(Message(message,
                                         position=self.position,
                                         standard=self.standard,
                                         message_type=message_type
                                         ))
        else:
            self.messages.append(Message(message,
                                         position=position,
                                         standard=self.standard,
                                         message_type=message_type
                                         ))

        self.ok = False

    def __add__(self, other):
        if self.position and not other.position:
            for i in range(len(other.messages)):
                if not other.messages[i].position:
                    other.messages[i].position = self.position
                if not other.messages[i].standard and self.standard:
                    other.messages[i].standard = self.standard

        self.messages += other.messages
        if not other.ok:
            self.ok = False
        return self
