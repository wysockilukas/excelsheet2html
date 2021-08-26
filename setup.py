import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="excelsheet2html",
    version="0.0.1",
    author="Lukasz Wysocki",
    description="Excelsheet2html excel table to html",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
    # py_modules=["excelsheet2html"],
    # package_dir={'': 'excelsheet2html'},
    install_requires=['openpyxl>=2.4.8']
)
