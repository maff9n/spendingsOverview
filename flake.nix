{
  description = "Google Apps Script development environment";

  inputs.nixpkgs.url = "github:nixos/nixpkgs/nixpkgs-unstable";

  outputs = { self, nixpkgs }:
    let
      system = "x86_64-linux";
      pkgs = import nixpkgs { inherit system; };
    in
    {
      devShells.${system}.default =
        pkgs.mkShell {
          packages = with pkgs; [
            google-clasp
            nodejs
          ];
        };
    };
}
