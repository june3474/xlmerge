# -*- coding: utf-8 -*-

import sys
import os
from string import Template
from xlwings import App


class Addin:
    def __init__(self):
        """Class for creating xlmerge addin.

        The addin, of which name is ``xlmerge.xlam``, contains several macros
        which runs ``xlmerge`` or performs merging.
        The addin is copied in '%UserProfile%\AppData\Roaming\Microsoft\AddIns'
        directory by the installer.

        """

        self.inst_dir = os.path.dirname(os.path.realpath(__file__))
        self.template = self.read_template()

    def read_template(self):
        """Read ``addin.template`` file"""

        # Determine the current path
        template = os.path.join(self.inst_dir, 'addin.template')
        with open(template, encoding='utf-8') as f:
            return Template(f.read())

    def insert_macro(self, wb, exe_path):
        """Insert necessary macros in the given workbook.

        As a sidenote, instead of adding from string, it is possible to import a saved module whole.
        wb.api.VBProject.VBComponents.Import(os.path.join(os.getcwd(), 'xlmerge.bas'))

        Args:
            wb (:obj:`xlwings.Book`): Workbook into which macros be inserted
            exe_path (str): full path of `xlmerge.exe`

        """

        macro = self.template.substitute(xlmergePath=exe_path)
        # '1' stands for standard module
        module = wb.api.VBProject.VBComponents.Add(1)
        module.name = 'xlmerge'
        module.CodeModule.AddFromString(macro)

    def generate(self, xlmergePath):
        """Generate Excel adddin ``xlmerge.xlam``

        The addin is saved in the same directory as ``addin.exe``

        Args:
            xlmergePath: Full path of ``xlmerge.exe``.
            This path substitutes the placeholder in ``addin.template`` file.

        """

        exe_path = os.path.join(xlmergePath, 'xlmerge.exe')
        with App(visible=False) as excel:
            wb = excel.books.active
            self.insert_macro(wb, exe_path)
            wb.save('xlmerge.xlam')
            wb.close()


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(1)

    addin = Addin()
    addin.generate(sys.argv[1])
