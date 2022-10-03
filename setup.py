from setuptools import setup, find_packages

setup(
    name='utilities_counter_importer',
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'utilities_counter_importer=utilities_counter_importer.command_line:main'
        ]
    },
    install_requires=[
        'google_cloud'
    ]
)
