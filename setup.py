from distutils.core import setup
import py2exe, PIL

setup(console=["main.py"])
setup(windows=[{'script': 'main.py'}],
      options={"py2exe": {"bundle_files": 1, "includes": ["PIL"]}})
# setup(windows=[{'script': 'main.py'}],
#       options={"py2exe": {"bundle_files": 3, 'packages': ['PIL']}},
#       zipfile="lib/library.zip")
