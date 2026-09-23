def main() -> None:
    import uvicorn

    uvicorn.run("ked_builder_web.app:app", host="0.0.0.0", port=8000, reload=False)


if __name__ == "__main__":
    main()
