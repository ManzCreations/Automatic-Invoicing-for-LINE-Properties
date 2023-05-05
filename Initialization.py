# Create a catch to make sure certain packages are installed if need be.
import pip


def install(app, package):
    pip.main(['install', package])
    app.log(f'Package "{package}" successfully installed!')
