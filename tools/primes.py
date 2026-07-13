"""
Hàm tính số nguyên tố
Ngày: 2026-03-31
"""

def is_prime(n):
    """Kiểm tra n có phải số nguyên tố không."""
    if n < 2:
        return False
    if n == 2:
        return True
    if n % 2 == 0:
        return False
    for i in range(3, int(n**0.5) + 1, 2):
        if n % i == 0:
            return False
    return True


def primes_up_to(limit):
    """Trả về danh sách số nguyên tố từ 2 đến limit."""
    return [n for n in range(2, limit + 1) if is_prime(n)]


if __name__ == "__main__":
    result = primes_up_to(100)
    print(f"Số nguyên tố từ 2 đến 100:")
    print(result)
    print(f"Tổng cộng: {len(result)} số nguyên tố")