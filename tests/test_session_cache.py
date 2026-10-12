"""Regression tests for safe per-chat file cache; runs with stdlib unittest."""

from __future__ import annotations

import unittest
from unittest.mock import patch

from session_cache import SessionFileCache, build_file_context, make_scope, mentions_file


class Event:
    def __init__(self, group="", user="u1", platform="qq_official", adapter="bot1"):
        self.group = group
        self.user = user
        self.platform = platform
        self.adapter = adapter

    def get_platform_name(self):
        return self.platform

    def get_platform_id(self):
        return self.adapter

    def get_group_id(self):
        return self.group

    def get_sender_id(self):
        return self.user


class CacheTests(unittest.TestCase):
    def test_same_group_shared_but_other_groups_isolated(self):
        cache = SessionFileCache()
        group_a = make_scope(Event(group="group-A"))
        same_group_member = make_scope(Event(group="group-A", user="u2"))
        group_b = make_scope(Event(group="group-B"))
        other_bot = make_scope(Event(group="group-A", adapter="bot2"))
        self.assertEqual(group_a, same_group_member)
        self.assertNotEqual(group_a, group_b)
        self.assertNotEqual(group_a, other_bot)
        cache.add(group_a, filename="report.txt", digest="one", text="hello", truncated=False)
        self.assertEqual(len(cache.list(same_group_member)), 1)
        self.assertEqual(cache.list(group_b), [])
        self.assertEqual(cache.list(other_bot), [])

    def test_private_and_platform_isolation(self):
        a = make_scope(Event(user="u1"))
        b = make_scope(Event(user="u2"))
        c = make_scope(Event(user="u1", platform="weixin_oc"))
        self.assertNotEqual(a, b)
        self.assertNotEqual(a, c)

    def test_duplicate_and_file_count(self):
        cache = SessionFileCache(max_files_per_scope=2)
        self.assertTrue(cache.add("group", filename="a", digest="x", text="first", truncated=False))
        self.assertFalse(cache.add("group", filename="a", digest="x", text="first", truncated=False))
        cache.add("group", filename="b", digest="y", text="second", truncated=False)
        cache.add("group", filename="c", digest="z", text="third", truncated=False)
        self.assertEqual([item.filename for item in cache.list("group")], ["b", "c"])

    def test_ttl(self):
        cache = SessionFileCache(lifetime_seconds=60)
        cache.add("g", filename="a", digest="a", text="hello", truncated=False, now=10)
        with patch("session_cache.time.monotonic", return_value=69):
            self.assertEqual(len(cache.list("g")), 1)
        with patch("session_cache.time.monotonic", return_value=70):
            self.assertEqual(cache.list("g"), [])

    def test_global_char_budget(self):
        cache = SessionFileCache(max_total_chars=1000)
        cache.add("a", filename="a", digest="a", text="a" * 700, truncated=False, now=10)
        cache.add("b", filename="b", digest="b", text="b" * 700, truncated=False, now=20)
        with patch("session_cache.time.monotonic", return_value=25):
            self.assertFalse(cache.list("a"))
            self.assertEqual(len(cache.list("b")), 1)

    def test_ambiguous_and_explicit_selection(self):
        cache = SessionFileCache()
        cache.add("g", filename="A.pdf", digest="a", text="alpha", truncated=False)
        cache.add("g", filename="B.pdf", digest="b", text="beta", truncated=False)
        selected, ambiguous = cache.choose("g", "帮我看看文件")
        self.assertEqual(selected, [])
        self.assertTrue(ambiguous)
        selected, ambiguous = cache.choose("g", "分析 A.pdf")
        self.assertEqual([d.filename for d in selected], ["A.pdf"])
        self.assertFalse(ambiguous)
        selected, ambiguous = cache.choose("g", "分析最新文件")
        self.assertEqual([d.filename for d in selected], ["B.pdf"])
        self.assertFalse(ambiguous)

    def test_no_secret_in_unrelated_prompt(self):
        self.assertFalse(mentions_file("今天天气怎么样", ["report.pdf"]))
        self.assertTrue(mentions_file("总结 report.pdf", ["report.pdf"]))
        cache = SessionFileCache()
        cache.add("g", filename="a", digest="a", text="SECRET", truncated=True)
        output = build_file_context(cache.list("g"))
        self.assertIn("SECRET", output)
        self.assertIn("不完整", output)


if __name__ == "__main__":
    unittest.main()
