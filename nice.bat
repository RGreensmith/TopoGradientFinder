docker build -t findslopeimg .
docker run --name topoGradientFinder -it findslopeimg python3 thing
docker rm topoGradientFinder --force