namespace ExceLint
    module BasicStats =
        
        /// <summary>Indicator function.</summary>
        /// <param name="b">index</param>
        /// <param name="t">threshold</param>
        /// <param name="xs">input vector</param>
        /// <returns>1 if true, 0 if false</returns>
        let indicator(b: int)(t: double)(xs: double[]) : double =
            if xs.[b] < t then 1.0 else 0.0

        /// <summary>Given a threshold value, computes the proportion of values < t.</summary>
        /// <param name="t">threshold</param>
        /// <param name="xs">input vector</param>
        /// <returns>a proportion between 0 and 1</returns>
        let cdf(t: double)(xs: double[]) : double =
            let inds = Array.map (fun i -> indicator i t xs) [| 0 .. (xs.Length - 1) |]
            (Array.sum inds) / System.Convert.ToDouble(xs.Length)

        /// <summary>For a given p value, find the value of x in X such that p <= Pr(x).</summary>
        /// <param name="p">p cumulative probability</param>
        /// <param name="xs">input vector</param>
        /// <returns>a value of x in X</returns>
        let cdfInverse(p: double)(xs: double[]) : double =
            assert (p >= 0.0)
            assert (p <= 1.0)

            // sort and zip with index
            let xs' = Array.sort xs

            // find the first index where cumPr > 
            let idx_opt = Array.tryFindIndex (fun x ->
                            (cdf x xs') > p
                          ) xs'

            match idx_opt with
            // if an index is found, take the one before
            | Some i -> if i = 0 then xs'.[0] else xs'.[i - 1]
            // if an index is not found, take the max element
            | None   -> xs'.[xs'.Length - 1]

        let mean(X: double[]) : double =
            X |> Array.sum |> (fun acc -> acc / double X.Length)

        let median(X: double[]) : double =
            let X' = Array.sort X
            if X'.Length % 2 = 0 then
                mean([|X'.[X'.Length / 2]; X'.[X'.Length / 2 - 1]|])
            else
                X'.[X'.Length / 2]

        let counts<'a when 'a : comparison>(X: 'a[]) : int[] =
            let d = new Utils.Dict<'a, int>()

            // count values of X
            for x in X do
                if d.ContainsKey(x) then
                    d.[x] <- d.[x] + 1
                else
                    d.Add(x, 1)

            d.Values |> Seq.toArray

        // find the multinomial probability vector for a sample;
        // probabilties are in value-sorted order
        let empiricalProbabilities(Y: int[]) : double[] =
            // compute probabilities
            let n = double (Array.sum Y)
            Y
            |> Seq.toArray
            |> Array.sort
            |> Array.map (fun count -> double count / n)

        /// <summary>
        /// Returns the entropy for the given multinomial probability vector.
        /// </summary>
        /// <param name="P">Vector of probabilities, one for each outcome of the random variable X.</param>
        let entropy(P: double[]) : double =
            let products = P |> Array.map (fun p -> p * System.Math.Log(p, 2.0))
            let sum = Array.sum products
            let result = -sum
            result

        /// <summary>
        /// Returns the normalized entropy (aka the "efficiency") for the given
        /// multinomial probability vector.
        /// </summary>
        /// <param name="P"></param>
        let normalizedEntropy(P: double[])(n: int) : double =
            if n = 1 then
                0.0
            else
                let products = P |> Array.map (fun p -> p * System.Math.Log(p, 2.0))
                let sum = Array.sum products
                let maxEntropy = System.Math.Log(double n, 2.0)
                let result = -sum / maxEntropy
                result