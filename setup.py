import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="cover-gen",
    version="3.3.0",
    author="Sunny Son",
    author_email="sunnys2327@gmail.com",
    description="Generates cover letter from application template",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3.7",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires = '>=3.7',
    install_requires=[
        "openpyxl",
        "docx2pdf",
        "docxtpl",
        "pandas",
        "python-dateutil"
    ],
)