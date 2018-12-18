echo Uploading all changes
# find . -size +30M | cat >> .gitignore
git add -A
git commit -a
git push
git ftp push