import { Switch, Route } from "wouter";
import Navbar from "@/components/Navbar";
import CatalogMapper from "@/pages/CatalogMapper";
import BgColorTweaker from "@/pages/BgColorTweaker";
import NotFound from "@/pages/not-found";

function App() {
  return (
    <div className="min-h-screen flex flex-col bg-[#1a1a2e]">
      <Navbar />
      <div className="flex-1">
        <Switch>
          <Route path="/" component={CatalogMapper} />
          <Route path="/bg-tweaker" component={BgColorTweaker} />
          <Route component={NotFound} />
        </Switch>
      </div>
    </div>
  );
}

export default App;
