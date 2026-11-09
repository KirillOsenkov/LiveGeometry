@echo off
rem Serves the whole site from this working tree at http://localhost:5012, nothing outside the
rem repo needed: the Browser head is published first (Debug, no AOT, into its bin/site: about
rem a minute the first time, some ten seconds after), which gives the app, the player and the
rem gallery as they are now, and the repo's static folders are laid over it, so that edits to
rem the Geometry Snacks and Catriona Agg libraries, /web, /embed and /history show on a reload
rem (open pages reload themselves), their puzzle pages answered as on the site. A change to
rem the app or the player shows on the next start. Ctrl+C stops it.
cd /d "%~dp0"
dotnet publish Main/Avalonia/LiveGeometry.Browser/LiveGeometry.Browser.csproj -c Debug -p:RunAOTCompilation=false -o Main/Avalonia/LiveGeometry.Browser/bin/site --nologo -v q || exit /b 1
echo Live Geometry at http://localhost:5012  (Ctrl+C stops the server)
dotnet run tools/serve.cs -- Main/Avalonia/LiveGeometry.Browser/bin/site/wwwroot 5012 --overlay snacks=snacks --route snacks=snacks/puzzle.html --overlay CatrionaAgg=CatrionaAgg --route CatrionaAgg=CatrionaAgg/puzzle.html --overlay web=web --overlay embed=embed --overlay history=history --reload
