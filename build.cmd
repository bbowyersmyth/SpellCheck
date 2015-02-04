msbuild src\SpellCheck.sln /t:clean /p:Configuration=Release /p:Platform="Any CPU"
msbuild src\SpellCheck.sln /p:Configuration=Release /p:Platform="Any CPU" /p:VisualStudioVersion=12.0
md builds
md builds\local
.\.NuGet\nuget.exe pack NuSpec\SpellCheck.nuspec -OutputDirectory builds\local
