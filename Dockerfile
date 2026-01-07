FROM public.ecr.aws/shelf/lambda-libreoffice-base:25.2-python3.13-x86_64

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY lambda_function.py /var/task/

CMD ["lambda_function.lambda_handler"]
