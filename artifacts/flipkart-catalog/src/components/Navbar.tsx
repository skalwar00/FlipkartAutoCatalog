import { useLocation, Link } from "wouter";

const pages = [
  { path: "/",        label: "Catalog Mapper",       icon: "📋" },
  { path: "/bg-tweaker", label: "BG Color Tweaker",  icon: "🎨" },
];

export default function Navbar() {
  const [location] = useLocation();

  return (
    <nav className="w-full bg-[#12122a] border-b border-white/10 px-6 py-0 flex items-center gap-1 sticky top-0 z-50 shadow-lg">
      <span className="text-[#f9a825] font-bold text-base tracking-tight mr-6 py-3 select-none">
        ⚡ FlipTools
      </span>

      {pages.map(({ path, label, icon }) => {
        const active = location === path;
        return (
          <Link key={path} href={path}>
            <span
              className={`
                flex items-center gap-2 px-4 py-3 text-sm font-medium cursor-pointer
                border-b-2 transition-colors duration-150 select-none
                ${active
                  ? "border-[#f9a825] text-[#f9a825]"
                  : "border-transparent text-white/50 hover:text-white/90 hover:border-white/30"}
              `}
            >
              <span>{icon}</span>
              {label}
            </span>
          </Link>
        );
      })}
    </nav>
  );
}
