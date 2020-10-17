#!/bin/sh

CONTAINER_NAME=slibu

docker container run -i --rm -u $(id -u):$(id -g) -v "$(pwd)":/data:z -w /data "$CONTAINER_NAME" slibu

