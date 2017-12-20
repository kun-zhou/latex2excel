from distutils.core import setup

setup(
    name='latex2excel',
    version='0.1dev',
    description='Convert LaTeX tables to Excel sheets',
    keywords='latex excel',
    author='Kun Zhou',
    url='https://github.com/kun-zhou/latex2excel',
    license='GNU GPLv3',
    py_modules=['latex2excel'],
    install_requires=['openpyxl','Click'],
     entry_points='''
        [console_scripts]
        latex2excel=latex2excel:main
    ''',
)
