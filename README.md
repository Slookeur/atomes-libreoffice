# atomes extension for [LibreOffice][libreoffice]

![License][license]
![Development Status][dev_status]

Extension to insert [atomes][atomes] projet file(s) in any [LibreOffice][libreoffice] document. 

## Install instructions

Download the file [atomes_extension.oxt][atomes_extension] that contains the extension.

### Using [LibreOffice][libreoffice]

Open [LibreOffice][libreoffice], press `Ctrl+Alt+E` (to open the extension manager), then browse to select the file `atomes_extension.oxt` and install. 

### Using the terminal

```[bash]
unopkg add atomes_extension.oxt
```

## Remove instructions

### Using the terminal

```[bash]
unopkg remove atomes_extension.oxt
```
 
## Build instructions

To build (package) the extension: 

```
./build.sh
```

## Dependency for Linux

On Linux Python scripts support is not included in the base installation of [LibreOffice][libreoffice], 
therefore you need to install it: 

### Debian based Linux:

```[bash]
sudo apt install libreoffice-script-provider-python
```

### RedHat based Linux:

```[bash]
sudo dnf install libreoffice-pyuno
```

## Beta version

This work is still in the testing stage, for the time being [atomes][atomes] project files cannot be stored in the ODF. 


## Who's behind ***atomes*** (and this extension)


***atomes*** is developed by [Dr. Sébastien Le Roux][slr], research engineer for the [CNRS][cnrs]

<p align="center">
  <a href="https://www.cnrs.fr/"><img width="100" src="https://upload.wikimedia.org/wikipedia/fr/thumb/7/72/Logo_Centre_national_de_la_recherche_scientifique_%282023-%29.svg/langfr-250px-Logo_Centre_national_de_la_recherche_scientifique_%282023-%29.svg.png" alt="CNRS logo" align="center"></a>
</p>

[Dr. Sébastien Le Roux][slr] works at the Institut de Physique et Chimie des Matériaux de Strasbourg [IPCMS][ipcms]

<p align="center">
  <a href="https://www.ipcms.fr/"><img width="100" src="https://www.ipcms.fr/uploads/2020/09/cropped-dessin_logo_IPCMS_couleur_vectoriel_r%C3%A9%C3%A9quilibr%C3%A9-2.png" alt="IPCMS logo" align="center"></a>
</p>

[atomes_extension]:https://github.com/Slookeur/atomes-libreoffice/tree/main/atomes_extensio.oxt
[libreoffice]:https://www.libreoffice.org/
[license]:https://img.shields.io/badge/License-AGPL_v3%2B-blue
[dev_status]:https://www.repostatus.org/badges/latest/active.svg
[slr]:https://www.ipcms.fr/sebastien-le-roux/
[cnrs]:https://www.cnrs.fr/
[ipcms]:https://www.ipcms.fr/
[github]:https://github.com/
[jekyll]:https://jekyllrb.com/
[atomes]:https://atomes.ipcms.fr/
[atomes-doc]:https://slookeur.github.io/atomes-doc/
[atomes-tuto]:https://slookeur.github.io/atomes-tuto/
[dlpoly]:https://www.scd.stfc.ac.uk/Pages/DL_POLY.aspx
[lammps]:https://lammps.sandia.gov/
[cpmd]:http://www.cpmd.org
[cp2k]:http://cp2k.berlios.de
[gtk]:https://www.gtk.org/
[openmp]:https://www.openmp.org/
