#!/bin/sh

git pull
cd Depends; git checkout master; git pull
cd parcel; git checkout master; git pull
cd ../ParcelCOMShim; git checkout master; git pull
cd parcel; git checkout master; git pull

