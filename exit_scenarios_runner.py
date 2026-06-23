#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Standalone runner for exit-scenarios cache.
Runs in a separate process so Telegram bot cannot hang if calculations are slow.
"""
import argparse
import importlib.util
import json
import os
import sys
import traceback


def _load_bot_module(bot_file: str):
    bot_file = os.path.abspath(bot_file)
    spec = importlib.util.spec_from_file_location("masif_bot_for_exit_cache", bot_file)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load bot module from {bot_file}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--bot-file", required=True)
    p.add_argument("--cache", default="exit_probabilities_cache.json")
    args = p.parse_args()
    try:
        mod = _load_bot_module(args.bot_file)
        fn = getattr(mod, "_v48_update_exit_probabilities_all", None)
        if not callable(fn):
            raise RuntimeError("_v48_update_exit_probabilities_all not found in bot module")
        # ensure cache path inside bot module points to requested path
        try:
            setattr(mod, "V48_EXIT_PROB_CACHE_FILE", args.cache)
        except Exception:
            pass
        data = fn()
        if not isinstance(data, dict):
            raise RuntimeError("exit update returned non-dict data")
        print(json.dumps({"ok": True, "updated_at": data.get("updated_at"), "count": data.get("count", 0)}, ensure_ascii=False))
        return 0
    except Exception as e:
        print(json.dumps({"ok": False, "error": f"{type(e).__name__}: {str(e)[:500]}", "trace": traceback.format_exc()[-1200:]}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
