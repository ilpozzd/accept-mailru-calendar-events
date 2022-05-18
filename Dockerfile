FROM python:3.7-alpine as base

ARG GID=1000
ARG UID=1000

RUN addgroup -g ${GID} cgroup && \
    adduser -u ${UID} -G cgroup -D cuser

RUN apk add --update --no-cache \
    build-base openldap-dev python3-dev && \
    echo -n "INPUT ( libldap.so )" > /usr/lib/libldap_r.so

FROM base as requirements

USER cuser:cgroup
WORKDIR /home/cuser

COPY requirements.txt .

RUN export PATH="/home/cuser/.local/bin:$PATH" && \
    python -m pip install --upgrade pip && \
    pip install -r requirements.txt

FROM requirements as code

WORKDIR /home/cuser/app
COPY --chown=cuser:cgroup ./app .

CMD ["python", "app.py"]