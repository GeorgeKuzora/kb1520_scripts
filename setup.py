from setuptools import setup

setup(name='main_prog',
    version='0.1',
    description='KB1520 scripts',
    url='https://github.com/GeorgeKuzora/kb1520_scripts',
    author='Georgiy Kuzora',
    author_email='rafale87@gmail.com',
    license='GPL',
    packages=['main_prog'],
    zip_safe=False,
    install_requires=[
    'et-xmlfile==1.0.1',
    'numpy==1.18.4',
    'openpyxl==3.0.7',
    'pandas==1.0.3',
    'python-dateutil==2.8.1',
    'pytz==2021.1',
    'six==1.15.0',
    'xlrd==2.0.1',
    'xlwt==1.3.0',
    ]
)
