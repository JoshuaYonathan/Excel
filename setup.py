from setuptools import setup

setup(name='excel',
      version='0.1',
      description='Utilizes windows excel com objects to control excel applications',
      url='https://github.com/JoshuaYonathan/Excel',
      author='Joshua Yonathan',
      author_email='JoshYonathan@outlook.com',
      license='GNU GPL v3.0',
      packages=['excel'],
      install_requires=[
        'pywin32',
      ],
      zip_safe=False)
