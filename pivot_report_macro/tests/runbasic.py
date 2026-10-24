"""Запуск Basic-кода в LibreOffice (headless) через UNO.

Код ядра макроса выполняется в LibreOffice Basic в режиме VBASupport —
это тот же язык, что VBA, без объектной модели Excel.
"""
import os, subprocess, tempfile, time
import uno
from com.sun.star.beans import PropertyValue


class Basic:
    def __init__(self, port=2299):
        self.prof = tempfile.mkdtemp(prefix='lo_profile_')
        self.p = subprocess.Popen(
            ['soffice', '--headless', '--invisible', '--norestore',
             '-env:UserInstallation=file://' + self.prof,
             f'--accept=socket,host=localhost,port={port};urp;'],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext(
            'com.sun.star.bridge.UnoUrlResolver', local)
        for _ in range(120):
            try:
                self.ctx = res.resolve(
                    f'uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext')
                break
            except Exception:
                time.sleep(0.5)
        self.desk = self.ctx.ServiceManager.createInstanceWithContext(
            'com.sun.star.frame.Desktop', self.ctx)
        pv = PropertyValue(); pv.Name = 'Hidden'; pv.Value = True
        self.doc = self.desk.loadComponentFromURL('private:factory/scalc', '_blank', 0, (pv,))

    def load(self, src, module='M'):
        libs = self.doc.BasicLibraries
        lib = libs.getByName('Standard') if libs.hasByName('Standard') else libs.createLibrary('Standard')
        if lib.hasByName(module):
            lib.removeByName(module)
        lib.insertByName(module, src)
        self.module = module

    def call(self, entry, *args):
        sp = self.doc.getScriptProvider()
        s = sp.getScript(f'vnd.sun.star.script:Standard.{self.module}.{entry}'
                         '?language=Basic&location=document')
        return s.invoke(args, (), ())[0]

    def close(self):
        try:
            self.doc.close(True)
            self.desk.terminate()
        except Exception:
            pass
        self.p.wait(timeout=30)
