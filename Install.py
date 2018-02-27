import pip


def install(package):
    pip.main(['install', package])


# Example
if __name__ == '__main__':
    install('pyttk')
    install('python-tk')
    install('google-api-python-client')
    install('python-dateutil')
    install('pillow')
    install('matplotlib')
    install('numpy')
    install('pyinstaller')
