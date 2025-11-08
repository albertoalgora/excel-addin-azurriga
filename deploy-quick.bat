@echo off
echo Building production version...
call npm run build

echo Copying manifest to dist...
copy /Y manifest-production.xml dist\manifest.xml

echo Committing changes...
git add .
git commit -m "Quick deploy: %date% %time%"

echo Pushing to main...
git push origin main

echo Deploying to GitHub Pages...
git subtree push --prefix dist origin gh-pages

echo.
echo ============================================
echo Deployment complete!
echo Your changes are live at:
echo https://albertoalgora.github.io/excel-addin-azurriga/
echo ============================================
pause
