// Numbers as the library writes them: Math.Round(value, digits) of Math.cs (halves away from
// zero, through decimal, so that 2.675 rounds as written), double.ToString() (the shortest
// text that reads back), ToString("F2") and ToStringInvariant.

const NumberFormat = {
    /**
     * Math.Round(num, fractionalDigits): rounded for showing, a half up as at school (3.125
     * is 3.13), the number taken as it is written (2.675 as a double is a hair under, and
     * would otherwise round down). Plus zero: no negative zero.
     */
    round(value, fractionalDigits) {
        if (!isValidValue(value)) {
            return value;
        }

        if (Math.abs(value) >= 1e15) {
            return roundToDigits(value, fractionalDigits) + 0;
        }

        // the number as decimal sees it: its 15 significant digits
        const sign = value < 0 ? -1 : 1;
        const text = Math.abs(value).toExponential(14);
        const match = /^(\d)\.(\d+)e([+-]\d+)$/.exec(text);
        if (match == null) {
            return roundToDigits(value, fractionalDigits) + 0;
        }

        let digits = match[1] + match[2];
        const exponent = parseInt(match[3], 10);

        // the position of the decimal point in digits, then the digit to round at
        const pointAt = exponent + 1;
        const keep = pointAt + fractionalDigits;
        if (keep < 0) {
            return 0;
        }

        if (keep >= digits.length) {
            return sign * Math.abs(value) + 0;
        }

        let kept = keep === 0 ? "0" : digits.substring(0, keep);
        const next = digits.charCodeAt(keep) - 48;
        if (next >= 5) {
            kept = NumberFormat.incrementDigits(kept);
        }

        // kept is an integer in units of 10^-fractionalDigits
        const scaled = Number(kept);
        const result = scaled / Math.pow(10, fractionalDigits);
        return sign * result + 0;
    },

    incrementDigits(digits) {
        const array = digits.split("");
        let i = array.length - 1;
        while (i >= 0) {
            if (array[i] === "9") {
                array[i] = "0";
                i--;
            } else {
                array[i] = String.fromCharCode(array[i].charCodeAt(0) + 1);
                return array.join("");
            }
        }

        return "1" + array.join("");
    },

    /** double.ToString(): the shortest text that reads back as the same number */
    toString(value) {
        if (Number.isNaN(value)) {
            return "NaN";
        }

        if (value === Infinity) {
            return "Infinity";
        }

        if (value === -Infinity) {
            return "-Infinity";
        }

        const text = String(value);
        // .NET writes 1E-05 where JavaScript writes 0.00001; the exponent forms differ too,
        // but no label shows such numbers rounded to two decimals
        return text === "-0" ? "0" : text;
    },

    /** ToString("F" + digits): every decimal, zeros included */
    toFixed(value, digits) {
        if (!isValidValue(value)) {
            return NumberFormat.toString(value);
        }

        const text = value.toFixed(digits);
        return /^-0(\.0*)?$/.test(text) ? text.substring(1) : text;
    },

    /** ToDegreeString: "{0:F0}°" */
    toDegreeString(degrees) {
        return NumberFormat.toFixed(degrees, 0) + "°";
    },

    /** ToRadianString: "{0:F02} rad" */
    toRadianString(radians) {
        return NumberFormat.toFixed(radians, 2) + " rad";
    }
};
