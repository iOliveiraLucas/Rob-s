FROM ubuntu:latest
RUN apt-get -y update
RUN apt-get install -y apt-utils
RUN apt-get install -y python3-pip python3-dev build-essential
RUN cd /usr/local/bin \
  && ln -s /usr/bin/python3 python \
  && pip3 install --upgrade pip
RUN apt-get install -y unzip curl
ENV TZ=America/Sao_Paulo
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone
# RUN apt-get install -y default-jdk 
RUN curl -sS -o - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add
RUN echo "deb [arch=amd64]  http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list
RUN apt-get -y update
RUN apt-get -y install google-chrome-stable
RUN pip install mysql-connector-python
COPY . /python
WORKDIR /python
RUN pip install -r requirements.txt
CMD ["python","-u","app.py"]