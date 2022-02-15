# -*- coding: utf-8 -*-

# for python 2 & 3 compatibility
try:
    import mock  # First try python 2.7.x
except ImportError:
    from unittest import mock
