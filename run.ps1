# Shortcut: build and run the interactive Google Maps parser
$IMAGE = "google-maps-parser:latest"
New-Item -ItemType Directory -Force -Path output | Out-Null

Write-Host "Building Docker image..." -ForegroundColor Cyan
docker build -t $IMAGE .

docker run -it --rm -v "${PWD}\output:/output" -v "${PWD}\parser.py:/app/parser.py" $IMAGE