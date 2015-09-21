FROM alpine:3.2
MAINTAINER Caber Chu

RUN apk add --update \
	python \
	py-pip \
    && rm -rf /var/cache/apk/* \
    && pip install --upgrade \
    pip

RUN mkdir /app
WORKDIR /app
COPY requirements.txt /app/


RUN /usr/bin/pip install -r /app/requirements.txt

ENTRYPOINT ["python"]