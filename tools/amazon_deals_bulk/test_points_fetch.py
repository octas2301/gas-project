# -*- coding: utf-8 -*-
from points_fetch import points_from_listings_offer, percent_from_listings


def test_points_from_listings_offer():
    body = {
        "offers": [
            {
                "offerType": "B2B",
                "price": {"amount": "4435.0"},
            },
            {
                "offerType": "B2C",
                "price": {"amount": "4480.0"},
                "points": {"pointsNumber": 45},
            },
        ]
    }
    pct, yen = points_from_listings_offer(body)
    assert pct == 1, pct
    assert yen == 45, yen
    assert percent_from_listings(body) == 1


def test_points_zero():
    body = {
        "offers": [
            {"offerType": "B2C", "price": {"amount": "1000"}, "points": {"pointsNumber": 0}}
        ]
    }
    pct, yen = points_from_listings_offer(body)
    assert pct == 0 and yen == 0


if __name__ == "__main__":
    test_points_from_listings_offer()
    test_points_zero()
    print("ALL OK")
