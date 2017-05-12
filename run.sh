
if [ ! -d "node_modules" ]; then
  npm install
fi
node index $1
