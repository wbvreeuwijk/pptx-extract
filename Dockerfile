FROM alpine

LABEL version="1.0"
LABEL description="Docker image for converting PPTX files to QVX"
LABEL maintainer="bvk@qlik.com"

RUN apk update \
 && apk add ca-certificates wget openssl openjdk8

RUN wget https://github.com/wbvreeuwijk/pptx-extract/releases/download/0.1/pptx-extract.zip \
 && unzip -o pptx-extract.zip \
 && rm  pptx-extract.zip

RUN mkdir /config \
 && mkdir /presentations \
 && mkdir /data

ENV GOOGLE_APPLICATION_CREDENTIALS=/config/google.json

VOLUME /config
VOLUME /presentations
VOLUME /data

# Insert schedule in crontab
RUN echo "0 0,3,9,12,15,18,21 * * *  /pptx-extract/bin/pptx-extract /presentations /data" >> /etc/crontabs/root

WORKDIR /config
#CMD ["/qs-google-translate/bin/google-translate-server"]
CMD ["crond","-f","-d","8"]
