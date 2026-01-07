FROM public.ecr.aws/lambda/python:3.10

RUN yum install -y libxml2 libxslt && yum clean all

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY lambda_function.py .

CMD ["lambda_function.lambda_handler"]
