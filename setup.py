from setuptools import setup, find_packages # type: ignore

setup(
    name='nom_du_projet',
    version='1.0',
    description='Description de votre projet',
    author='Votre nom',
    author_email='votre@email.com',
    packages=find_packages(),
    install_requires=[
        'black==21.9b0',
        'matplotlib==3.4.3',
        'pandas==1.3.3',
        'streamlit==1.4.0',
        'toml==0.10.2',
        'xlwings==0.24.10',
    ],
)
