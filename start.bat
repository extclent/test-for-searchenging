@echo off
start cmd /k python "server.py"
timeout /t 3
start cmd /k python "start_tunnel.py"