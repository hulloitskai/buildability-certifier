const createWindowsInstaller = require('electron-winstaller').createWindowsInstaller;
const path = require('path');

getInstallerConfig()
     .then(createWindowsInstaller)
     .catch((error) => {
     console.error(error.message || error)
     process.exit(1)
 })

function getInstallerConfig () {
    console.log('creating windows installer')
    const rootPath = path.join('./')
    const outPath = path.join(rootPath, 'release-builds')

    return Promise.resolve({
       appDirectory: path.join(outPath, 'Buildability\ Certifier-win32-ia32/'),
       authors: 'Steven Xie',
       noMsi: true,
       outputDirectory: path.join(outPath, 'windows-installer'),
       exe: 'Buildability Certifier.exe',
       setupExe: 'Buildability Certifier Setup.exe',
       setupIcon: path.join(rootPath, 'production_assets', 'icons', 'win', 'icon.ico')
   });
}
