import aiohttp
import asyncio

API_URL = "https://unstandardisable-shantelle-gnawable.ngrok-free.dev/api/process/combinedOcrBytes"


async def process_file(file, customer_name, model_name):

    try:
        data = aiohttp.FormData()

        data.add_field(
            "files",
            file.getvalue(),
            filename=file.name,
            content_type="application/pdf"
        )

        data.add_field(
            "customer_name",
            customer_name
        )
        data.add_field("model_name", model_name)

        async with aiohttp.ClientSession() as session:
            async with session.post(
                API_URL,
                data=data
            ) as response:

                return await response.json()

    except Exception as e:
        return {"error": str(e)}
