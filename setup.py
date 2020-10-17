from setuptools import setup, find_packages

setup(
    name='slibu',
    version='0.1',
    # py_modules=['slibu'],
    packages=find_packages(include=['slibu']),
    install_requires=[
        'Click',
        'mistune==2.0.0a2',
        'Pillow',
        'Pygments',
        'python-pptx',
    ],
    entry_points='''
        [console_scripts]
        slibu=slibu.cli:build
    ''',
)
