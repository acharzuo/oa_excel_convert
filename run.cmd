@echo off
if (not exist node_modules)  npm install
npm start %1
open %1