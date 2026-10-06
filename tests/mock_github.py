"""GitHub API GIẢ LẬP (tối thiểu) để kiểm thử gh_store: contents (liệt kê / raw), git blobs/trees/commits/refs."""
import base64, hashlib, json, threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs


class Kho:
    def __init__(self):
        self.blobs, self.trees, self.commits = {}, {}, {}
        self.khoa = threading.Lock()
        t = self._tree({})
        c = self._commit(t, [], "init")
        self.head = c

    def _sha(self, b):
        return hashlib.sha1(b).hexdigest()

    def blob(self, data):
        s = hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()
        self.blobs[s] = data
        return s

    def _tree(self, m):
        s = self._sha(json.dumps(m, sort_keys=True).encode())
        self.trees[s] = dict(m)
        return s

    def _commit(self, tree, parents, msg):
        s = self._sha(json.dumps([tree, parents, msg, len(self.commits)]).encode())
        self.commits[s] = {"tree": tree, "parents": parents, "message": msg}
        return s


K = None


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, obj=None, raw=None):
        self.send_response(code)
        self.end_headers()
        if raw is not None:
            self.wfile.write(raw)
        elif obj is not None:
            self.wfile.write(json.dumps(obj).encode())

    def _body(self):
        n = int(self.headers.get("content-length", 0))
        return json.loads(self.rfile.read(n) or b"{}")

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        parts = u.path.strip("/").split("/")  # repos/o/r/...
        rest = parts[3:]
        with K.khoa:
            if rest[:3] == ["git", "ref", "heads"]:
                return self._send(200, {"object": {"sha": K.head}})
            if rest[:2] == ["git", "commits"]:
                c = K.commits[rest[2]]
                return self._send(200, {"tree": {"sha": c["tree"]}})
            if rest[:2] == ["git", "blobs"]:
                return self._send(200, raw=K.blobs[rest[2]])
            if rest[:1] == ["contents"]:
                ref = q.get("ref", ["main"])[0]
                cm = K.commits.get(ref) or K.commits[K.head]
                tree = K.trees[cm["tree"]]
                p = "/".join(rest[1:])
                if p in tree:
                    return self._send(200, raw=K.blobs[tree[p]])
                ds = [{"name": k[len(p) + 1:], "sha": v, "type": "file"} for k, v in tree.items()
                      if k.startswith(p + "/") and "/" not in k[len(p) + 1:]]
                return self._send(200 if ds else 404, ds if ds else {"message": "Not Found"})
            if u.path.endswith("/rate_limit"):
                return self._send(200, {"resources": {"core": {"remaining": 4999, "limit": 5000}}})
            return self._send(200, {"full_name": "o/r", "private": True, "permissions": {"push": True}})

    def do_POST(self):
        b = self._body()
        rest = urlparse(self.path).path.strip("/").split("/")[3:]
        with K.khoa:
            if rest == ["git", "blobs"]:
                return self._send(201, {"sha": K.blob(base64.b64decode(b["content"]))})
            if rest == ["git", "trees"]:
                m = dict(K.trees[b["base_tree"]])
                for it in b["tree"]:
                    if it["sha"] is None:
                        m.pop(it["path"], None)
                    else:
                        m[it["path"]] = it["sha"]
                return self._send(201, {"sha": K._tree(m)})
            if rest == ["git", "commits"]:
                return self._send(201, {"sha": K._commit(b["tree"], b["parents"], b["message"])})

    def do_PATCH(self):
        b = self._body()
        with K.khoa:
            c = K.commits[b["sha"]]
            if c["parents"] != [K.head]:
                return self._send(422, {"message": "Update is not a fast forward"})
            K.head = b["sha"]
            return self._send(200, {"object": {"sha": K.head}})


def chay(port=0):
    global K
    K = Kho()
    srv = ThreadingHTTPServer(("127.0.0.1", port), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv
