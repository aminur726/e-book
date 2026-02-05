#!/bin/bash

# -------------------------------
# App Config
# -------------------------------
AppVersion="1.0.0"
DockerHubUser="aminur726"
DockerHubRepoName="e-book"
DockerHubRepository="${DockerHubUser}/${DockerHubRepoName}"

# -------------------------------
# DockerHub Login
# -------------------------------
docker login --username ${DockerHubUser} --password 2024bdfuneelbuilder

# -------------------------------
# Backend Service (Streamlit App)
# -------------------------------
BackendService="ebook-app-service"
BackendServiceDir="."

echo "Building Docker Image for ${BackendService}"

# -------------------------------
# Build Multi-Arch Image (AMD64 + ARM64)
# -------------------------------
docker buildx create --use --name mybuilder || true

docker buildx build \
  --platform linux/amd64,linux/arm64 \
  --no-cache \
  -f ${BackendServiceDir}/Dockerfile \
  -t ${BackendService}:${AppVersion} \
  -t ${DockerHubRepository}:${BackendService}-${AppVersion} \
  --push \
  ${BackendServiceDir}

echo "✅ Build & Push Complete!"
echo "📌 Pushed Image: ${DockerHubRepository}:${BackendService}-${AppVersion}"
