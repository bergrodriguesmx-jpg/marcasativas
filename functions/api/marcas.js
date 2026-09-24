export async function onRequest(context) {
  try {
    const { results } = await context.env.DB.prepare("SELECT * FROM marcas").all();

    return new Response(JSON.stringify(results), {
      headers: { 
        "content-type": "application/json;charset=UTF-8",
        "access-control-allow-origin": "*"
      }
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500,
      headers: { "content-type": "application/json" }
    });
  }
}
