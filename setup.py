from setuptools import setup, find_packages

setup(
    name="google_doc_backup",
    version="0.1.0",
    description="A tool to backup Google Docs/Sheets/Slides to Office formats.",
    author="Eric Cochran",
    author_email="ecochran76@gmail.com",
    url="https://github.com/ecochran76/google_doc_backup",  # if applicable
    packages=find_packages(),
    install_requires=[
        "pydrive",
        "python-dateutil",
    ],
    entry_points={
        "console_scripts": [
            "google-doc-backup=google_doc_backup.cli:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)
