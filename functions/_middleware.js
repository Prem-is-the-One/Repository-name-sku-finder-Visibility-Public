function unauthorized() {
    return new Response("Authentication required.", {
        status: 401,
        headers: {
            "WWW-Authenticate": 'Basic realm="SKU Finder", charset="UTF-8"',
            "Cache-Control": "no-store"
        }
    });
}

export async function onRequest(context) {

    const authHeader =
        context.request.headers.get("Authorization");

    if (!authHeader || !authHeader.startsWith("Basic ")) {
        return unauthorized();
    }

    try {

        const encoded =
            authHeader.substring(6);

        const decoded =
            atob(encoded);

        const separator =
            decoded.indexOf(":");

        if (separator === -1) {
            return unauthorized();
        }

        const username =
            decoded.substring(0, separator);

        const password =
            decoded.substring(separator + 1);

        const correctUser =
            context.env.AUTH_USER;

        const correctPassword =
            context.env.AUTH_PASS;

        if (
            username !== correctUser ||
            password !== correctPassword
        ) {
            return unauthorized();
        }

        return context.next();

    } catch (error) {

        return unauthorized();

    }
}