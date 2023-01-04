from setuptools import setup

setup(
    name='google_sheets',
    version='1.0.0b1',
    author='Vitor Gabriel',
    author_email='edvitor13@hotmail.com',
    packages=['google_sheets'],
    python_requires='>=3.10',
    install_requires=[
        'pydantic^=1.10.2',
        'google-api-python-client^=2.65.0',
        'google-auth-httplib2^=0.1.0',
        'google-auth-oauthlib^=0.7.1'
    ],
)
