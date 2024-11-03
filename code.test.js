const code = require('./code');

test('adds 1 + 2 to equal 3', () => {
  expect(code.plus(1, 2)).toBe(3);
});

