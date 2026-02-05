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
docker login --password 2024bdfuneelbuilder --username ${DockerHubUser}

# -------------------------------
# Backend Service (Streamlit App)
# -------------------------------
BackendService="ebook-app-service"
BackendServiceDir="."

echo "Building Docker Image for ${BackendService}"

# Build Image
docker image build --no-cache -f ${BackendServiceDir}/Dockerfile \
  -t ${BackendService}:${AppVersion} ${BackendServiceDir}

# Tag Image
docker image tag ${BackendService}:${AppVersion} \
  ${DockerHubRepository}:${BackendService}-${AppVersion}

# Push Image to DockerHub
echo "Pushing Image to DockerHub..."
docker push ${DockerHubRepository}:${BackendService}-${AppVersion}

echo "✅ Build & Push Complete!"
echo "📌 Pushed Image: ${DockerHubRepository}:${BackendService}-${AppVersion}"

### End-Of-File ###
