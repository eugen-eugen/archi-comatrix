FROM alpine:latest

COPY dist/ /archi-scripts/

CMD ["/bin/sh"]
