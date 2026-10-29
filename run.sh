#!/bin/bash
# Build and run the interactive Google Maps parser
set -e
export MSYS_NO_PATHCONV=1
IMAGE="google-maps-parser:latest"
mkdir -p output

echo "Building image..."
docker build -t "$IMAGE" .

docker run -it --rm -v "${PWD}/output:/output" -v "${PWD}/parser.py:/app/parser.py" "$IMAGE"
