FROM golang:1.19

WORKDIR /app

COPY go.mod go.sum ./

RUN go mod download

COPY *.go ./

RUN CGO_ENABLED=1 GOOS=linux go build -o /docker-whatsappLogger

CMD [ "/docker-whatsappLogger" ]
