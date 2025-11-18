FROM ubuntu:18.04

# Ubuntu libreoffice stuff is copied from https://github.com/ipunkt/docker-libreoffice-headless/blob/master/Dockerfile


RUN apt-get update && \
	apt-get -y -q install \
        vim \
        libxml2-dev libxslt-dev python-dev libjpeg-dev \
        python3 \
        python3-pip \
        openjdk-11-jre \
		imagemagick \
		libreoffice \
		libreoffice-writer \
		ure \
		libreoffice-java-common \
		libreoffice-core \
		libreoffice-common \
		#openjdk-8-jre \
		fonts-opensymbol \
		hyphen-fr \
		hyphen-de \
		hyphen-en-us \
		hyphen-it \
		hyphen-ru \
		fonts-dejavu \
		fonts-dejavu-core \
		fonts-dejavu-extra \
		fonts-droid-fallback \
		fonts-dustin \
		fonts-f500 \
		fonts-fanwood \
		fonts-freefont-ttf \
		fonts-liberation \
		fonts-lmodern \
		fonts-lyx \
		fonts-sil-gentium \
		fonts-texgyre \
		fonts-tlwg-purisa && \
	apt-get -y -q remove libreoffice-gnome && \
	apt -y autoremove && \
	rm -rf /var/lib/apt/lists/* \
	rm  /etc/ImageMagick-6/policy.xml 

RUN adduser --home=/opt/libreoffice --disabled-password --gecos "" --shell=/bin/bash libreoffice



WORKDIR /app

COPY pptx2md.py  /app

# Install required Python packages (uv runner, python-pptx and PyMuPDF)
RUN pip3 install uv python-pptx PyMuPDF

VOLUME /data
WORKDIR /data

ENTRYPOINT ["uv", "/app/pptx2md.py"]
