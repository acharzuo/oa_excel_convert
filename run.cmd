;;@echo off
%~d0
cd %~p0
if not exist node_modules ( npm install )
node index %1
pause