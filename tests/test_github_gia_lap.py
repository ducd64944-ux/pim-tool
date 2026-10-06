"""gh_store với GitHub giả lập: 2 máy chủ (2 Store) ghi đồng thời — không mất dữ liệu, xung đột được báo."""
import os, sys, threading
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gh_store, mock_github
from gh_store import KHONG_CO, Store


def test_hai_may():
    srv = mock_github.chay()
    gh_store.API = f"http://127.0.0.1:{srv.server_port}"
    A = Store("github", repo="o/r", token="x")
    B = Store("github", repo="o/r", token="x")
    assert A.luu({"shared/q.json": b'{"a":1}'}, "A1", {"shared/q.json": KHONG_CO})[0]
    da, pa = A.doc2("shared/q.json")
    db, pb = B.doc2("shared/q.json")
    assert da == db == b'{"a":1}' and pa == pb
    assert B.luu({"shared/q.json": b'{"a":1,"b":2}'}, "B", {"shared/q.json": pb})[0]
    ok, _, _, xd = A.luu({"shared/q.json": b'{"a":1,"c":3}'}, "A2", {"shared/q.json": pa})
    assert not ok and "shared/q.json" in xd                      # A bị chặn, không ghi đè B
    assert A.doc("shared/q.json") == b'{"a":1,"b":2}'
    # 20 luồng ghi 20 file KHÁC nhau cùng lúc -> đủ cả 20 (thử lại khi nhánh đổi)
    kq = []
    def ghi(i):
        S = A if i % 2 else B
        kq.append(S.luu({f"users/u{i}/settings.json": str(i).encode()}, f"t{i}", {f"users/u{i}/settings.json": KHONG_CO})[0])
    ts = [threading.Thread(target=ghi, args=(i,)) for i in range(20)]
    [t.start() for t in ts]
    [t.join() for t in ts]
    assert all(kq), kq
    C = Store("github", repo="o/r", token="x")
    assert all(C.doc(f"users/u{i}/settings.json") == str(i).encode() for i in range(20))
    assert C.doc("shared/q.json") == b'{"a":1,"b":2}'
    srv.shutdown()


if __name__ == "__main__":
    test_hai_may()
    print("OK — GitHub giả lập: ghi đồng thời không mất dữ liệu")
