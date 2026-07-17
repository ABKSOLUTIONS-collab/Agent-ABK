FROM agent365registry.azurecr.io/librechat:latest
COPY patch.js /tmp/patch.js
COPY abk-logo.png /app/client/dist/assets/abk-logo.png
COPY librechat.yaml /app/librechat.yaml
RUN node /tmp/patch.js
