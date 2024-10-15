from setuptools import setup, find_packages

setup(
    name="trello-card-creator",
    version="0.1.0",
    packages=find_packages(exclude=["tests"]),
    install_requires=[
        "python-docx>=0.8.11",
        "requests>=2.26.0",
        "keyring>=23.5.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.1.2",
        ],
    },
    entry_points={
        "console_scripts": [
            "trello-card-creator=trello_card_creator.trello_card_creator:main",
        ],
    },
)
