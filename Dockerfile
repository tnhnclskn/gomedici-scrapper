FROM python

WORKDIR /app
RUN mkdir export
ADD . .
RUN pip install -r requirements.txt

CMD python -u scrapper.py