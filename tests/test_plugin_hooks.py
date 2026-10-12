"""Black-box tests for AstrBot event integration, without installing AstrBot."""

from __future__ import annotations

import asyncio
import importlib
import sys
import tempfile
import types
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


class FileSegment:
    def __init__(self, path: str, name: str):
        self.path = path
        self.name = name

    async def get_file(self):
        return self.path


class FakeEvent:
    def __init__(self, group: str, user: str, files=None, *, adapter="qq1"):
        self.group = group
        self.user = user
        self.adapter = adapter
        self.files = files or []
        self.is_at_or_wake_command = False

    def get_platform_name(self):
        return "qq_official"

    def get_platform_id(self):
        return self.adapter

    def get_group_id(self):
        return self.group

    def get_sender_id(self):
        return self.user

    def get_messages(self):
        return self.files

    def plain_result(self, message):
        return message


def _load_plugin():
    class DummyFilter:
        EventMessageType = SimpleNamespace(ALL=1)

        @staticmethod
        def event_message_type(*args):
            return lambda handler: handler

        @staticmethod
        def command(*args):
            return lambda handler: handler

        @staticmethod
        def on_llm_request(*args):
            return lambda handler: handler

    class Star:
        def __init__(self, context):
            self.context = context

    class TextPart:
        def __init__(self, text=""):
            self.text = text
            self.temporary = False

        def mark_as_temp(self):
            self.temporary = True
            return self

    pkg_names = [
        "astrbot",
        "astrbot.api",
        "astrbot.api.event",
        "astrbot.api.star",
        "astrbot.api.provider",
        "astrbot.api.message_components",
        "astrbot.api.all",
        "astrbot.core",
        "astrbot.core.agent",
        "astrbot.core.agent.message",
    ]
    modules = {name: types.ModuleType(name) for name in pkg_names}
    for package in ("astrbot", "astrbot.api", "astrbot.core", "astrbot.core.agent"):
        modules[package].__path__ = []
    modules["astrbot.api.event"].filter = DummyFilter
    modules["astrbot.api.event"].AstrMessageEvent = type("AstrMessageEvent", (), {})
    modules["astrbot.api.star"].Context = type("Context", (), {})
    modules["astrbot.api.star"].Star = Star
    modules["astrbot.api.star"].register = lambda *a, **kw: lambda cls: cls
    modules["astrbot.api.provider"].ProviderRequest = type("ProviderRequest", (), {})
    modules["astrbot.api.message_components"].File = FileSegment
    modules["astrbot.core.agent.message"].TextPart = TextPart
    modules["astrbot.api"].logger = SimpleNamespace(
        warning=lambda *a, **kw: None,
        exception=lambda *a, **kw: None,
    )
    root = Path(__file__).resolve().parents[1]
    pkg = types.ModuleType("_reader_pr_test")
    pkg.__path__ = [str(root)]
    modules["_reader_pr_test"] = pkg
    with patch.dict(sys.modules, modules):
        return importlib.import_module("_reader_pr_test.main")


class PluginTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.main = _load_plugin()

    def setUp(self):
        self.plugin = self.main.astrbot_plugin_file_reader(context=None)

    def test_receive_without_wake_and_inject_only_to_same_group(self):
        async def run():
            # Place test fixtures in the requested D: project workspace.
            with tempfile.TemporaryDirectory(
                dir=Path(__file__).resolve().parents[1]
            ) as tmp:
                f = Path(tmp, "report.txt")
                f.write_text("GROUP_A_PRIVATE_DATA", encoding="utf-8")
                sent = FakeEvent(
                    "group-A", "user-1", [FileSegment(str(f), "report.txt")]
                )
                await self.plugin.on_receive_msg(sent)
                self.assertEqual(
                    len(
                        self.plugin.cache.list("qq_official\x1fqq1\x1fgroup\x1fgroup-A")
                    ),
                    1,
                )
                req = SimpleNamespace(
                    prompt="请总结 report.txt", extra_user_content_parts=[]
                )
                await self.plugin.on_request(FakeEvent("group-A", "user-2"), req)
                self.assertEqual(len(req.extra_user_content_parts), 1)
                self.assertIn(
                    "GROUP_A_PRIVATE_DATA", req.extra_user_content_parts[0].text
                )
                self.assertTrue(req.extra_user_content_parts[0].temporary)

                other = SimpleNamespace(
                    prompt="请总结 report.txt", extra_user_content_parts=[]
                )
                await self.plugin.on_request(FakeEvent("group-B", "user-3"), other)
                self.assertEqual(other.extra_user_content_parts, [])

                other_adapter = SimpleNamespace(
                    prompt="请总结 report.txt", extra_user_content_parts=[]
                )
                await self.plugin.on_request(
                    FakeEvent("group-A", "user-3", adapter="qq2"), other_adapter
                )
                self.assertEqual(other_adapter.extra_user_content_parts, [])

                casual = SimpleNamespace(prompt="你好", extra_user_content_parts=[])
                await self.plugin.on_request(FakeEvent("group-A", "user-3"), casual)
                self.assertEqual(casual.extra_user_content_parts, [])

                await self.plugin.on_receive_msg(sent)
                self.assertEqual(
                    len(
                        self.plugin.cache.list("qq_official\x1fqq1\x1fgroup\x1fgroup-A")
                    ),
                    1,
                )

        asyncio.run(run())

    def test_scanlike_empty_text_rejected(self):
        async def run():
            with tempfile.TemporaryDirectory(
                dir=Path(__file__).resolve().parents[1]
            ) as tmp:
                f = Path(tmp, "blank.txt")
                f.write_text("", encoding="utf-8")
                await self.plugin.on_receive_msg(
                    FakeEvent("group-A", "user-1", [FileSegment(str(f), "blank.txt")])
                )
                self.assertEqual(self.plugin.reading_history, [])

        asyncio.run(run())


if __name__ == "__main__":
    unittest.main()
