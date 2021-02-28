FROM python:3.9

WORKDIR /usr/app/prj

RUN apt-get update \
    && apt-get install -y locales \
    && locale-gen ja_JP.UTF-8
ENV LANG ja_JP.UTF-8
ENV LANGUAGE ja_JP:ja
ENV LC_ALL=ja_JP.UTF-8
RUN localedef -f UTF-8 -i ja_JP ja_JP.utf8

COPY requirements.txt ./
RUN pip install -r requirements.txt

COPY . .

EXPOSE 5000
#CMD ["python", "server.py"]

# ==============================================================================
# build
# sudo docker build -t my-flusk-app .

# ==============================================================================
# run
# sudo docker run -it --network host -v /home/aw/work_dir/06_project:/usr/app/lifelog --name flusk-app -d my-flusk-app