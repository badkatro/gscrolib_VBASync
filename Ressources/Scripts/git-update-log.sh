touch temp.log
git log -1 --date=format:'%Y-%m-%d %H:%M:%S' | grep Date -A 2 | grep -Ev "^$" | sed -e 's/^[ ]*//' | cat - update.log > temp.log
cat temp.log | tee update.log
rm temp.log