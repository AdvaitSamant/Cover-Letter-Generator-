import reflex as rx


config = rx.Config(
    app_name="cover_letter",
    plugins=[
        rx.plugins.SitemapPlugin(),
        rx.plugins.RadixThemesPlugin(),
    ],
)
