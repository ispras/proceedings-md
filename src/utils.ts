
export function *uniqueNameGenerator(name: string) {
    let index = 0
    while (true) {
        let nameCandidate = name
        if (index > 0) nameCandidate += "_" + index
        yield nameCandidate
        index++;
    }
}