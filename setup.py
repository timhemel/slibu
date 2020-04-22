from setuptools import setup

setup(
    name='slibu',
    version='0.1',
    py_modules=['slibu'],
    install_requires=[
        'Click',
        'mistune==2.0.0a2',
        'Pillow',
        'Pygments',
        'python-pptx',
    ],
    entry_points='''
        [console_scripts]
        slibu=slibu:cli
    ''',
)
