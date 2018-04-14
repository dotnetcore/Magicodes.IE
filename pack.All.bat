call ./clear.bat
call ./pack.bat "Magicodes.ExporterAndImporter.Core*.nupkg" "./src/Magicodes.ExporterAndImporter.Core/Magicodes.ExporterAndImporter.Core.csproj"

call ./pack.bat "Magicodes.ExporterAndImporter.Excel*.nupkg" "./src/Magicodes.ExporterAndImporter.Excel/Magicodes.ExporterAndImporter.Excel.csproj"
@pause