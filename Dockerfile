FROM public.ecr.aws/lambda/python:3.13

COPY requirements.txt .
RUN pip install -r requirements.txt --target "${LAMBDA_TASK_ROOT}"

COPY gmail/ ${LAMBDA_TASK_ROOT}/gmail/
COPY processors/ ${LAMBDA_TASK_ROOT}/processors/
COPY storage/ ${LAMBDA_TASK_ROOT}/storage/
COPY utils/ ${LAMBDA_TASK_ROOT}/utils/
COPY config.py ${LAMBDA_TASK_ROOT}
COPY lambda_function.py ${LAMBDA_TASK_ROOT}

CMD ["lambda_function.lambda_handler"]