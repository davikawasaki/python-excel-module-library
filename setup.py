import setuptools


def read_dependencies(req_file):
    with open(req_file) as req:
        return [line.strip() for line in req]


with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="pythonexcel",
    version="0.2.5",
    author="Davi Kawasaki",
    author_email="davishinjik@gmail.com",
    description="Python Excel Modules",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/davikawasaki/python-excel-module-library",
    packages=setuptools.find_packages(),
    install_requires=read_dependencies("requirements.txt")
)