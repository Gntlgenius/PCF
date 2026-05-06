# 1. Scaffold fresh project
pac pcf init --namespace acc --name ClientFlagViewer --template field

# 2. Replace generated files with the ones above

# 3. Install + regenerate types
npm install
npm run refreshTypes

# 4. Build & push
npm run build
pac pcf push --publisher-prefix acc
