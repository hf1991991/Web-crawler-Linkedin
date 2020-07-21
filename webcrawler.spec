# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['/Users/henriquefalconer/Desktop/Poli Júnior/NTec /Projetos/Webcrawler Linkedin/python_files/macos_main.py'],
             pathex=['/Users/henriquefalconer/Desktop/Poli Júnior/NTec /Projetos/Webcrawler Linkedin'],
             binaries=[],
             datas=[('/Users/henriquefalconer/Desktop/Poli Júnior/NTec /Projetos/Webcrawler Linkedin/python_files/webcrawler', 'webcrawler/')],
             hiddenimports=['pkg_resources.py2_warn', 'scrapy.spiderloader', 'scrapy.statscollectors', 'scrapy.logformatter', 'scrapy.extensions', 'scrapy.extensions.logstats', 'scrapy.extensions.corestats', 'scrapy.extensions.memusage', 'scrapy.extensions.feedexport', 'scrapy.extensions.memdebug', 'scrapy.extensions.closespider', 'scrapy.extensions.throttle', 'scrapy.extensions.telnet', 'scrapy.extensions.spiderstate', 'scrapy.core.scheduler', 'scrapy.core.downloader', 'scrapy.downloadermiddlewares', 'scrapy.downloadermiddlewares.robotstxt', 'scrapy.downloadermiddlewares.httpauth', 'scrapy.downloadermiddlewares.downloadtimeout', 'scrapy.downloadermiddlewares.defaultheaders', 'scrapy.downloadermiddlewares.useragent', 'scrapy.downloadermiddlewares.retry', 'scrapy.core.downloader.handlers.http', 'scrapy.core.downloader.handlers.s3', 'scrapy.core.downloader.handlers.ftp', 'scrapy.core.downloader.handlers.datauri', 'scrapy.core.downloader.handlers.file', 'scrapy.downloadermiddlewares.ajaxcrawl', 'scrapy.core.downloader.contextfactory', 'scrapy.downloadermiddlewares.redirect', 'scrapy.downloadermiddlewares.httpcompression', 'scrapy.downloadermiddlewares.cookies', 'scrapy.downloadermiddlewares.httpproxy', 'scrapy.downloadermiddlewares.stats', 'scrapy.downloadermiddlewares.httpcache', 'scrapy.spidermiddlewares', 'scrapy.spidermiddlewares.httperror', 'scrapy.spidermiddlewares.offsite', 'scrapy.spidermiddlewares.referer', 'scrapy.spidermiddlewares.urllength', 'scrapy.spidermiddlewares.depth', 'scrapy.pipelines', 'scrapy.dupefilters', 'queuelib', 'scrapy.squeues'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='webcrawler',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )
