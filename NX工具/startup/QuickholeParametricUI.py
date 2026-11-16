# -*- coding: utf-8 -*-
"""Unique NX UI entry; never resolve the older Quickhole.py by basename."""
import importlib.util
import os
import sys

import NXOpen


def load_dialog_module():
    name = 'codex_quickhole_parameter_ui'
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Quickhole.py')
    cached = sys.modules.get(name)
    if cached is not None:
        if os.path.normcase(os.path.abspath(cached.__file__)) != os.path.normcase(path):
            raise RuntimeError('快捷腰孔参数界面路径冲突，请完整重启 NX。')
        return cached
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(name, None)
        raise
    return module


def main():
    load_dialog_module().main()


def GetUnloadOption(dummy):
    return NXOpen.Session.LibraryUnloadOption.AtTermination


if __name__ == '__main__':
    main()
