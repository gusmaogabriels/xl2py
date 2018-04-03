from setuptools import setup


requires = ['numpy','pythoncom','win32com','re','copy','time']

packages = [
		'xl2py',
		'xl2py.core',
		'xl2py.conversion_lib',
		'xl2py.com_handlers',
]

package_dir = {'xl2py' : 'xl2py'}
package_data = { 'xl2py' : []}


setup(
    name='xl2py',
    version='1.0.2b',
    packages=packages,
    license='The MIT License (MIT)',
    author = 'Gabriel S. Gusmao',
    author_email = 'gusmaogabriels@gmail.com',
    url = 'https://github.com/gusmaogabriels/xl2py',
    download_url = 'https://github.com/gusmaogabriels/xl2py/tarball/v1.0.2b',
    keywords = ['python', 'excel', 'com', 'solver', 'optmization', 'minimization', 'evolutionary', 'stochastic'],
    package_data = package_data,
    package_dir = package_dir,

)
