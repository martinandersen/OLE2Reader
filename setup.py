from distutils.core import setup

setup(name='OLE2Reader',
    version='0.1',
    description='Module for importing data from OLE 2 files',
    author='Martin S. Andersen',
    author_email='martin.skovgaard.andersen@gmail.com',
    url='https://github.com/martinandersen/OLE2Reader',
    download_url='https://github.com/martinandersen/OLE2Reader/archive/master.zip',
    license = 'GNU GPL version 3',
    package_dir = {"": "python"},
    packages = [""],
    requires=["JPype1","Numpy"],
    classifiers=['Development Status :: 4 - Beta',
                 'Programming Language :: Python'])
