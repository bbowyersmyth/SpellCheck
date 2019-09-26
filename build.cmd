@echo off
msbuild src\SpellCheck.sln /t:clean /p:Configuration=Release /p:Platform="Any CPU" /m
msbuild src\SpellCheck.sln /p:Configuration=Release /p:Platform="Any CPU" /p:VisualStudioVersion=14.0 /m
if not exist builds md builds
if not exist builds\local md builds\local
copy src\SpellCheck\bin\Release\*.nupkg builds\local
