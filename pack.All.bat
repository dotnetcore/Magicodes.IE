call ./clear.bat
call ./pack.bat "Magicodes.IE.Core*.nupkg" "./src/Magicodes.ExporterAndImporter.Core/Magicodes.ExporterAndImporter.Core.csproj"

call ./pack.bat "Magicodes.IE.Excel*.nupkg" "./src/Magicodes.ExporterAndImporter.Excel/Magicodes.ExporterAndImporter.Excel.csproj"
@pause