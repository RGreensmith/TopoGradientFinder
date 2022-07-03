FROM qgis/qgis
WORKDIR /findSlope
COPY ./requirements.txt .
RUN pip install -r requirements.txt
COPY . .